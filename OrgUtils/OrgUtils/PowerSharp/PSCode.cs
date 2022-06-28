using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
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
        public PSCode(string computername, string code, CodeType codeType, Hashtable parameters, List<FileInfo> modulesTobeImported, uint timeoutInSeconds)
        {
            this.ComputerName = computername;
            this.Code = code;
            this.CodeType = codeType;
            this.Parameters = parameters != null ? CSharpConverters.HashtableToDictionary<string, object>(parameters) : null;
            this.ModulesTobeImported = modulesTobeImported;
            this.TimeoutInSeconds = timeoutInSeconds;
        }
        #endregion

        #region Public methods
        public PSExecutionResult Invoke()
        {
            CancellationTokenSource cts = new CancellationTokenSource();
            Task<PSExecutionResult> task = null;
            try
            {
                Func<PSExecutionResult> func = new Func<PSExecutionResult>(IgnitePowerShellInvocation);
                cts.CancelAfter(TimeSpan.FromSeconds(TimeoutInSeconds));
                task = Task.Run(func, cts.Token);
                task.Wait(cts.Token);
                return task.Result;
            }
            catch (OperationCanceledException)
            {
                return new PSExecutionResult()
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
                return new PSExecutionResult()
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
                cts.Dispose();
                if (task.Status == TaskStatus.RanToCompletion)
                    task.Dispose();
                task = null;
                cts = null;
            }
        }
        #endregion

        #region Private methods
        private PSExecutionResult IgnitePowerShellInvocation()
        {
            PSRuntimeEnvironment psRuntimeEnvironment = new PSRuntimeEnvironment(this);
            PSExecutionResult psExecutionResult_temp = new PSExecutionResult()
            {
                ComputerName = this.ComputerName,
                HadErrors = false,
                Results = null,
                Errors = null
            };
            try
            {
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
                psExecutionResult_temp = (PSExecutionResult)psExecutionResult_temp.Clone(); //Deep cloning
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
            finally
            {
                psRuntimeEnvironment.PowerShell.Dispose();
                psRuntimeEnvironment.Dispose();
                psRuntimeEnvironment = null;
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
                return RunspaceFactory.CreateRunspace(NewWSManConnection(this.ComputerName, openTimeoutInMilliseconds));
            }
        }
        private WSManConnectionInfo NewWSManConnection(string computerName, int sessionOpenTimeoutInMilliSeconds, int port = 5985)
        {
            return new WSManConnectionInfo(new Uri($"http://{computerName}:{port}/WSMAN")) { OpenTimeout = sessionOpenTimeoutInMilliSeconds };
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
        public uint TimeoutInSeconds { get; }
        public uint Throttle { get; }

        private Hashtable parmeters;
        Random lockObject = new Random();
        ParallelOptions po = new ParallelOptions();
        ObservableCollection<PSExecutionResult> results = new ObservableCollection<PSExecutionResult>();
        #endregion

        #region Constructor
        public PowerShell(List<string> computernames, string code, CodeType codeType, Hashtable parameters, List<FileInfo> modulesTobeImported, uint timeoutInSeconds = 30, uint throttle = 4)
        {
            this.ComputerNames = computernames;
            this.Code = code;
            this.CodeType = codeType;
            this.Parameters = parameters != null ? CSharpConverters.HashtableToDictionary<string, object>(parameters) : null;
            this.ModulesTobeImported = modulesTobeImported;
            this.TimeoutInSeconds = timeoutInSeconds;
            this.Throttle = throttle;

            this.parmeters = parameters;
        }
        #endregion

        #region Public methods
        public ObservableCollection<PSExecutionResult> BeginInvoke()
        {
            po.MaxDegreeOfParallelism = (int)this.Throttle;
            Parallel.ForEach<PSCode>(CreatePSCodeObjectsList(), po, psCode =>
            {
                var res = psCode.Invoke();
                lock (lockObject)
                {
                    results.Add(res);
                }
            });
            return results;
        }
        public void BeginInvoke(NotifyCollectionChangedEventHandler onEveryPSExecutionComplete)
        {
            po.MaxDegreeOfParallelism = (int)this.Throttle;
            results.CollectionChanged += onEveryPSExecutionComplete;
            Parallel.ForEach<PSCode>(CreatePSCodeObjectsList(), po, psCode =>
            {
                var res = psCode.Invoke();
                lock (lockObject)
                {
                    results.Add(res);
                }
            });
        }
        #endregion

        #region Private methods
        private IEnumerable<PSCode> CreatePSCodeObjectsList()
        {
            foreach (var computer in ComputerNames)
            {
                yield return (new PSCode(computer, this.Code, this.CodeType, this.parmeters, this.ModulesTobeImported, this.TimeoutInSeconds));
            }
        }
        #endregion
    }
}