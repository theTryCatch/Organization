using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
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
        public uint TimeoutInSeconds { get; }
        #endregion

        #region Constructor
        public PSCode(string computername, string code, CodeType codeType, Dictionary<string, object> parameters, List<FileInfo> modulesTobeImported, uint timeoutInSeconds)
        {
            this.ComputerName = computername;
            this.Code = code;
            this.CodeType = codeType;
            this.Parameters = parameters;
            this.ModulesTobeImported = modulesTobeImported;
            this.TimeoutInSeconds = timeoutInSeconds;
        }
        #endregion

        #region Public methods
        public PSExecutionResult Invoke()
        {
            PSExecutionResult result = null;
            using (CancellationTokenSource cts = new CancellationTokenSource())
            {
                Func<PSExecutionResult> func = new Func<PSExecutionResult>(IgnitePowerShellInvocation);
                cts.CancelAfter(TimeSpan.FromSeconds(TimeoutInSeconds));
                Task<PSExecutionResult> task = null;
                try
                {
                    task = Task.Run(func, cts.Token);
                    task.Wait(cts.Token);
                    result = task.Result;
                }
                catch (OperationCanceledException)
                {
                    result = new PSExecutionResult()
                    {
                        ComputerName = this.ComputerName,
                        HadErrors = true,
                        Errors = new PSDataCollection<ErrorRecord>() {
                            new ErrorRecord(
                                new Exception($"Execution timeout with in {this.TimeoutInSeconds} second(s)"),
                                string.Empty, ErrorCategory.OperationTimeout, null)
                        },
                        Results = null
                    };
                }
                catch (Exception e)
                {
                    result = new PSExecutionResult()
                    {
                        ComputerName = this.ComputerName,
                        HadErrors = true,
                        Errors = new PSDataCollection<ErrorRecord>() {
                            new ErrorRecord(
                                new Exception($"{e.InnerException.Message}"),
                                string.Empty, ErrorCategory.NotSpecified, null)
                        },
                        Results = null
                    };
                }
                finally
                {
                    // to handle this case we are not enclosing task in using blocks
                    if (task.Status == TaskStatus.RanToCompletion)
                        task.Dispose();
                }
            }
            return result;
        }
        #endregion

        #region Private methods
        private PSExecutionResult IgnitePowerShellInvocation()
        {
            PSExecutionResult psExecutionResult_temp = null;

            using (PSRuntimeEnvironment psRuntimeEnvironment = new PSRuntimeEnvironment(this))
            {
                try
                {
                    using (psRuntimeEnvironment.PowerShell)
                    {
                        var result = psRuntimeEnvironment.PowerShell.Invoke();

                        //If the executed code is just a command then the PowerShell.Invoke() method throws an exception but
                        //if the code is a script and contains any errors then it will not indicate anyway about the exceptions.
                        //Hence we the logic is around HadErrors property data.
                        if (psRuntimeEnvironment.PowerShell.HadErrors)
                        {
                            psExecutionResult_temp = new PSExecutionResult() { ComputerName = this.ComputerName, HadErrors = true, Results = null, Errors = psRuntimeEnvironment.PowerShell.Streams.Error };
                        }
                        else
                        {
                            psExecutionResult_temp = new PSExecutionResult() { ComputerName = this.ComputerName, HadErrors = false, Results = result, Errors = null };
                        }
                    }
                }
                catch (Exception e)
                {
                    psExecutionResult_temp = new PSExecutionResult()
                    {
                        ComputerName = this.ComputerName,
                        HadErrors = true,
                        Errors = new PSDataCollection<ErrorRecord>() { new ErrorRecord(e, string.Empty, ErrorCategory.OperationStopped, null) },
                        Results = null
                    };
                }
            }
            return psExecutionResult_temp;
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
                return RunspaceFactory.CreateRunspace(InitialSessionState.CreateDefault());
            }
            else
            {
                return RunspaceFactory.CreateRunspace(new WSManConnectionInfo(new Uri($"http://{this.ComputerName}:{5985}/WSMAN")) { OpenTimeout = openTimeoutInMilliseconds });
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

    public class PSExecutionResult
    {
        #region Public properties
        public string ComputerName { get; set; }
        public object Results { get; set; }
        public bool HadErrors { get; set; }
        public PSDataCollection<ErrorRecord> Errors { get; set; }
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
        public uint TimeoutInSeconds { get; }
        public int Throttle { get; }
        #endregion

        #region Constructor
        public PowerShell(List<string> computernames, string code, CodeType codeType, Dictionary<string, object> parameters, List<FileInfo> modulesTobeImported, uint timeoutInSeconds = 30, int throttle = 4)
        {
            this.ComputerNames = computernames;
            this.Code = code;
            this.CodeType = codeType;
            this.Parameters = parameters;
            this.ModulesTobeImported = modulesTobeImported;
            this.TimeoutInSeconds = timeoutInSeconds;
            this.Throttle = throttle;
        }
        #endregion

        #region Public methods
        public ObservableCollection<PSExecutionResult> BeginInvoke()
        {
            Random lockObject = new Random();
            ParallelOptions po = new ParallelOptions();
            ObservableCollection<PSExecutionResult> results = new ObservableCollection<PSExecutionResult>();
            po.MaxDegreeOfParallelism = this.Throttle;
            var psCodeObjects = CreatePSCodeObjectsList();
            Parallel.ForEach<PSCode>(psCodeObjects, po, psCode =>
            {
                var res = psCode.Invoke();
                lock (lockObject)
                {
                    results.Add(res);
                }
            });
            psCodeObjects = null;
            return results;
        }
        //public void BeginInvoke(NotifyCollectionChangedEventHandler onEveryPSExecutionComplete)
        //{
        //    Random lockObject = new Random();
        //    ParallelOptions po = new ParallelOptions();
        //    ObservableCollection<PSExecutionResult> results = new ObservableCollection<PSExecutionResult>();
        //    po.MaxDegreeOfParallelism = this.Throttle;
        //    results.CollectionChanged += onEveryPSExecutionComplete;
        //    Parallel.ForEach<PSCode>(CreatePSCodeObjectsList(), po, psCode =>
        //    {
        //        var res = psCode.Invoke();
        //        lock (lockObject)
        //        {
        //            results.Add(res);
        //        }
        //    });
        //}
        #endregion

        #region Private methods
        private IEnumerable<PSCode> CreatePSCodeObjectsList()
        {
            foreach (var computer in ComputerNames)
            {
                yield return (new PSCode(computer, this.Code, this.CodeType, this.Parameters, this.ModulesTobeImported, this.TimeoutInSeconds));
            }
        }
        #endregion
    }
    public class MyClass
    {
        public static void Main()
        {
            var a = new PowerSharp.PowerShell(new List<string>() { "mslaptop", "dc1", "asdlkfj", "mslaptop", "dc1", "asdlkfj", "dc1", "asdlkfj", "mslaptop", "dc1", "asdlkfj", "dc1", "asdlkfj", "mslaptop", "dc1", "asdlkfj", "dc1", "asdlkfj", "mslaptop", "dc1", "asdlkfj", "dc1", "asdlkfj", "mslaptop", "dc1", "asdlkfj", "dc1", "asdlkfj", "mslaptop", "dc1", "asdlkfj", "dc1", "asdlkfj", "mslaptop", "dc1", "asdlkfj", "dc1", "asdlkfj", "mslaptop", "dc1", "asdlkfj", "dc1", "asdlkfj", "mslaptop", "dc1", "asdlkfj", "dc1", "asdlkfj", "mslaptop", "dc1", "asdlkfj", "dc1", "asdlkfj", "mslaptop", "dc1", "asdlkfj", "dc1", "asdlkfj", "mslaptop", "dc1", "asdlkfj", "dc1", "asdlkfj", "mslaptop", "dc1", "asdlkfj", "dc1", "asdlkfj", "mslaptop", "dc1", "asdlkfj", "dc1", "asdlkfj", "mslaptop", "dc1", "asdlkfj", "dc1", "asdlkfj", "mslaptop", "dc1", "asdlkfj", "dc1", "asdlkfj", "mslaptop", "dc1", "asdlkfj", "dc1", "asdlkfj", "mslaptop", "dc1", "asdlkfj", "dc1", "asdlkfj", "mslaptop", "dc1", "asdlkfj", "dc1", "asdlkfj", "mslaptop", "dc1", "asdlkfj", "dc1", "asdlkfj", "mslaptop", "dc1", "asdlkfj", "dc1", "asdlkfj", "mslaptop", "dc1", "asdlkfj", "dc1", "asdlkfj", "mslaptop", "dc1", "asdlkfj", "dc1", "asdlkfj", "mslaptop", "dc1", "asdlkfj", "dc1", "asdlkfj", "mslaptop", "dc1", "asdlkfj", "dc1", "asdlkfj", "mslaptop", "dc1", "asdlkfj", "dc1", "asdlkfj", "mslaptop", "dc1", "asdlkfj", "dc1", "asdlkfj", "mslaptop", "dc1", "asdlkfj", "dc1", "asdlkfj", "mslaptop", "dc1", "asdlkfj", "dc1", "asdlkfj", "mslaptop", "dc1", "asdlkfj", "dc1", "asdlkfj", "mslaptop", "dc1", "asdlkfj", "dc1", "asdlkfj", "mslaptop", "dc1", "asdlkfj", "dc1", "asdlkfj", "mslaptop", "dc1", "asdlkfj", "dc1", "asdlkfj", "mslaptop", "dc1", "asdlkfj", "dc1", "asdlkfj", "mslaptop", "dc1", "asdlkfj", "dc1", "asdlkfj", "mslaptop", "dc1", "asdlkfj", "dc1", "asdlkfj", "mslaptop", "dc1", "asdlkfj", "dc1", "asdlkfj", "mslaptop", "dc1", "asdlkfj", "dc1", "asdlkfj", "mslaptop", "dc1", "asdlkfj", "dc1", "asdlkfj", "mslaptop", "dc1", "asdlkfj", "dc1", "asdlkfj", "mslaptop", "dc1", "asdlkfj", "dc1", "asdlkfj", "mslaptop", "dc1", "asdlkfj", "dc1", "asdlkfj", "mslaptop", "dc1", "asdlkfj", "dc1", "asdlkfj", "mslaptop", "dc1", "asdlkfj", "dc1", "asdlkfj", "mslaptop", "dc1", "asdlkfj", "dc1", "asdlkfj", "mslaptop", "dc1", "asdlkfj", "dc1", "asdlkfj", "mslaptop", "dc1", "asdlkfj" }, "get-service", CodeType.Cmdlet, null, null);
            var b = a.BeginInvoke();
            Console.Read();
        }
    }
}