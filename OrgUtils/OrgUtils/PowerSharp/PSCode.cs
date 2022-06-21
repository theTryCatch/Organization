using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Runtime.Serialization.Formatters.Binary;
using System.Threading;
using System.Threading.Tasks;
using MSFT = System.Management.Automation;
namespace Organization.PowerSharp
{
    public class PSCode
    {
        #region Public properties
        public string ComputerName { get; }
        public string Code { get; }
        public CodeType CodeType { get; }
        public Dictionary<string, object> Parameters { get; }
        public List<FileInfo> ModulesTobeImported { get; }
        #endregion

        #region Constructor
        public PSCode(string computername, string code, CodeType codeType, Hashtable parameters, List<FileInfo> modulesTobeImported)
        {
            this.ComputerName = computername;
            this.Code = code;
            this.CodeType = codeType;
            this.Parameters = parameters != null ? CSharpConverters.HashtableToDictionary<string, object>(parameters) : null;
            this.ModulesTobeImported = modulesTobeImported;
        }
        #endregion

        #region Public methods
        public PSExecutionResult Invoke()
        {
            PSExecutionResult psExecutionResult;
            try
            {
                using (var psRuntimeEnvironment = new PSRuntimeEnvironment(this))
                {
                    PSExecutionResult psExecutionResult_temp = new PSExecutionResult()
                    {
                        ComputerName = this.ComputerName,
                        HadErrors = false,
                        Results = null,
                        Errors = null
                    };
                    var result = psRuntimeEnvironment.PowerShell.Invoke();

                    //If the executed code is just a command then the PowerShell.Invoke() method throws an exception but
                    //if the code is a script and contains any errors then it will not indicate anyway about the exceptions.
                    //Hence we the logic is around HadErrors property data.
                    if (psRuntimeEnvironment.PowerShell.HadErrors)
                    {
                        psExecutionResult_temp.HadErrors = true;
                        psExecutionResult_temp.Errors = psRuntimeEnvironment.PowerShell.Streams.Error;
                        psExecutionResult_temp.Results = null;
                    }
                    else
                    {
                        psExecutionResult_temp.Results = result;
                    }
                    psExecutionResult = (PSExecutionResult)psExecutionResult_temp.Clone(); //Deep cloning
                }
            }
            catch (Exception e)
            {
                psExecutionResult = new PSExecutionResult()
                {
                    ComputerName = this.ComputerName,
                    HadErrors = true,
                    Errors = new PSDataCollection<ErrorRecord>() { new ErrorRecord(e, string.Empty, ErrorCategory.OperationStopped, null) },
                    Results = null
                };
            }
            return psExecutionResult;
        }
        public async Task<PSExecutionResult> InvokeAsync()
        {
            return await Task.Run(new Func<PSExecutionResult>(this.Invoke));
        }
        #endregion
    }

    public class PSRuntimeEnvironment : IDisposable
    {
        #region Public properties
        public string ComputerName { get; }
        public bool IsLocalComputer { get; }
        public MSFT.PowerShell PowerShell { get; }
        #endregion

        #region Constructor
        public PSRuntimeEnvironment(PSCode psCode)
        {
            this.ComputerName = psCode.ComputerName;
            this.IsLocalComputer = psCode.ComputerName.ToUpper() == Environment.MachineName.ToUpper() ? true : false;
            this.PowerShell = MSFT.PowerShell.Create();

            if (psCode.ModulesTobeImported != null)
            {
                foreach (var item in psCode.ModulesTobeImported)
                {
                    this.PowerShell.AddCommand("Import-Module").AddArgument(item.ToString());
                }
            }
            if (psCode.CodeType == CodeType.Cmdlet)
                this.PowerShell.AddCommand(psCode.Code);
            else if (psCode.CodeType == CodeType.Script)
                this.PowerShell.AddScript(psCode.Code);
            else
                this.PowerShell.AddScript(File.ReadAllText(psCode.Code));

            if (psCode.Parameters != null)
                this.PowerShell.AddParameters(psCode.Parameters);

            this.PowerShell.Runspace = NewRemoteRunspace();
            this.PowerShell.Runspace.Open();
        }
        #endregion

        #region Private methods
        private Runspace NewRemoteRunspace(int openTimeoutInMilliseconds = 5 * 1000)
        {
            if (this.IsLocalComputer)
            {
                var rrsp = RunspaceFactory.CreateRunspace(InitialSessionState.CreateDefault());
                return rrsp;
            }
            else
            {
                WSManConnectionInfo conInfo = NewWSManConnection(this.ComputerName, openTimeoutInMilliseconds);
                var rrsp = RunspaceFactory.CreateRunspace(conInfo);
                rrsp.InitialSessionState.Formats.Clear();
                return rrsp;
            }
        }
        private WSManConnectionInfo NewWSManConnection(string computerName, int sessionOpenTimeoutInMilliSeconds, int port = 5985)
        {
            Uri remoteComputerUri = new Uri($"http://{computerName}:{port}/WSMAN");
            var conInfo = new WSManConnectionInfo(remoteComputerUri);

            conInfo.OpenTimeout = sessionOpenTimeoutInMilliSeconds;
            return conInfo;
        }
        private void AddModulePaths(List<FileInfo> modulesTobeImported, ref Runspace runspace)
        {
            if (modulesTobeImported != null)
            {
                foreach (var path in modulesTobeImported)
                {
                    if (path != null)
                        if (File.Exists(path.ToString()))
                            runspace.InitialSessionState.ImportPSModulesFromPath(path.ToString());
                        else
                            throw new FileNotFoundException($"The file {path.ToString()} not found");
                }
            }
        }
        #endregion

        #region Interface implementations
        public void Dispose()
        {
            if (this.PowerShell.Runspace.RunspaceIsRemote)
                this.PowerShell.Runspace.Disconnect();
            this.PowerShell.Runspace.Close();
            this.PowerShell.Runspace.Dispose();
            this.PowerShell.Dispose();
        }
        #endregion
    }

    [Serializable]
    public class PSExecutionResult : ICloneable
    {
        #region Public properties
        public string ComputerName { get; set; }
        public object Results { get; set; }
        public bool HadErrors { get; set; }
        public PSDataCollection<ErrorRecord> Errors { get; set; }
        #endregion

        #region Interface implementations
        //Deep cloning
        public object Clone()
        {
            using (MemoryStream stream = new MemoryStream())
            {
                if (this.GetType().IsSerializable)
                {
                    BinaryFormatter formatter = new BinaryFormatter();
                    formatter.Serialize(stream, this);
                    stream.Position = 0;
                    return formatter.Deserialize(stream);
                }
                return null;
            }
        }
        #endregion
    }

    public class PowerShell
    {
        #region Public properties
        public List<string> ComputerNames { get; }
        public string Code { get; }
        public CodeType CodeType { get; }
        public Dictionary<string, object> Parameters { get; }
        public List<FileInfo> ModulesTobeImported { get; }

        private Hashtable parmeters;
        #endregion

        #region Constructor
        public PowerShell(List<string> computernames, string code, CodeType codeType, Hashtable parameters, List<FileInfo> modulesTobeImported)
        {
            this.ComputerNames = computernames;
            this.Code = code;
            this.CodeType = codeType;
            this.Parameters = parameters != null ? CSharpConverters.HashtableToDictionary<string, object>(parameters) : null;
            this.ModulesTobeImported = modulesTobeImported;

            this.parmeters = parameters;
        }
        #endregion

        #region Public methods
        public IEnumerable<PSExecutionResult> BeginInvoke()
        {
            CancellationTokenSource cts = new CancellationTokenSource();
            cts.CancelAfter(TimeSpan.FromSeconds(2));
            CancellationToken ctsToken = cts.Token;

            ParallelOptions po = new ParallelOptions();
            po.CancellationToken = ctsToken;
            po.MaxDegreeOfParallelism = Environment.ProcessorCount;

            ConcurrentBag<PSExecutionResult> psExecutionResults = new ConcurrentBag<PSExecutionResult>();
            try
            {
                Parallel.ForEach<PSCode>(CreatePSCodeObjectsList(), po, psCode =>
                  {
                      po.CancellationToken.ThrowIfCancellationRequested();
                      Thread.Sleep(5*1000);
                      //psExecutionResults.Add(psCode.Invoke());
                      Console.WriteLine("in the parallel loop");
                  });
            }
            catch (OperationCanceledException e)
            {
                Console.WriteLine(e.Message);
            }
            finally
            {
                cts.Dispose();
            }
            return psExecutionResults;
        }

        private IEnumerable<PSCode> CreatePSCodeObjectsList()
        {
            foreach (var computer in ComputerNames)
            {
                yield return (new PSCode(computer, this.Code, this.CodeType, this.parmeters, this.ModulesTobeImported));
            }
        }
        #endregion
    }

    public class Program
    {
        public static void Main()
        {
            var ps = new PowerSharp.PowerShell(new List<string>() { "mslaptop" }, "get-service; sleep 10", CodeType.Script, null, null);
            var b = ps.BeginInvoke();
            Console.WriteLine("All done");
            Console.ReadKey();
        }
    }
}