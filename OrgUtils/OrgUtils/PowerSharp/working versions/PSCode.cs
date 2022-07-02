//using Organization.PowerSharp;
//using System;
//using System.Collections.Generic;
//using System.Collections.ObjectModel;
//using System.Management.Automation;
//using System.Management.Automation.Runspaces;
//using System.Threading;
//using System.Threading.Tasks;
//using MSFT = System.Management.Automation;
//namespace OrgUtils.PowerSharp
//{
//    public class PSCode
//    {
//        #region Public properties
//        public Runspace Runspace { get; set; }
//        public MSFT.PowerShell PowerShell { get; } = MSFT.PowerShell.Create();
//        public string ComputerName { get; }
//        #endregion

//        #region Constructor
//        public PSCode(string computerName)
//        {
//            this.ComputerName = computerName;
//        }
//        #endregion

//        #region Public methods
//        public PSExecutionResult Invoke(
//            string script,
//            Dictionary<string, object> parameters,
//            List<string> modulesTobeImported,
//            byte timeoutInSeconds
//        )
//        {

//            #region importing the provided modules
//            //we are not using InitialSessionState or it's ImportPSModule method.
//            //for importing the provided modules keeping ...
//            //1. system resources in mind.
//            //2. This method will be silent if the provided path is invalid.

//            //Note: Import-Module should be on top of everything.


//            if (modulesTobeImported != null)
//            {
//                foreach (string module in modulesTobeImported)
//                {
//                    this.PowerShell.AddCommand("Import-Module").AddArgument(module);
//                }
//            }
//            #endregion

//            #region Adding script and it's parameters
//            if (!string.IsNullOrEmpty(script))
//                this.PowerShell.AddScript(script);
//            else
//                throw new ArgumentNullException("Script code can't be null or empty");

//            if (parameters != null)
//                this.PowerShell.AddParameters(parameters);

//            // else: parameters can be null
//            #endregion

//            try
//            {
//                AssignRunspace();
//                using (this.Runspace)
//                {
//                    OpenRunspace();
//                    #region Invoking the PowerShell
//                    using (CancellationTokenSource cts = new CancellationTokenSource())
//                    {
//                        cts.CancelAfter(TimeSpan.FromSeconds(timeoutInSeconds));
//                        Task<PSExecutionResult> task = null;

//                        Func<PSExecutionResult> psInvoke = new Func<PSExecutionResult>(IgnitePowerShellInvocation);
//                        try
//                        {
//                            task = Task.Run(psInvoke, cts.Token);
//                            task.Wait(cts.Token);
//                            return task.Result;
//                        }
//                        catch (OperationCanceledException)
//                        {
//                            return new PSExecutionResult()
//                            {
//                                ComputerName = this.ComputerName,
//                                HadErrors = true,
//                                Errors = new List<string>() { $"Execution timeout with in {timeoutInSeconds} second(s)" },
//                                Results = null
//                            };
//                        }
//                        finally
//                        {
//                            if (task.Status == TaskStatus.RanToCompletion || task.Status == TaskStatus.Canceled || task.Status == TaskStatus.Faulted)
//                                task.Dispose();
//                        }
//                    }
//                    #endregion
//                }
//            }
//            catch (Exception e)
//            {
//                return new PSExecutionResult()
//                {
//                    ComputerName = this.ComputerName,
//                    HadErrors = true,
//                    Errors = new List<string>() { e.ToString() },
//                    Results = null
//                };
//            }
//        }

//        #endregion

//        #region Private methods
//        private void OpenRunspace()
//        {
//            if (this.Runspace.RunspaceStateInfo.State == RunspaceState.BeforeOpen)
//                this.Runspace.Open();
//        }
//        private void AssignRunspace()
//        {
//            if (this.ComputerName != Environment.MachineName) // it is a local computer
//            {
//                this.Runspace = RunspaceFactory.CreateRunspace(new WSManConnectionInfo(new Uri($"http://{this.ComputerName}:{5985}/WSMAN")));
//                this.PowerShell.Runspace = this.Runspace;
//            }
//            else
//            {
//                this.Runspace = RunspaceFactory.CreateRunspace(InitialSessionState.CreateDefault());
//            }
//        }
//        private PSExecutionResult IgnitePowerShellInvocation()
//        {
//            using (this.PowerShell)
//            {
//                Collection<PSObject> result = null;
//                try
//                {
//                    result = this.PowerShell.Invoke();

//                    /*
//                     * If the executed code is just a command then the PowerShell.Invoke() method throws an exception but
//                    if the code is a script and contains any errors then it will not indicate anyway about the exceptions.
//                    Hence we the logic is around HadErrors property data.
//                    *
//                    */

//                    if (this.PowerShell.HadErrors)
//                    {
//                        List<string> errors = new List<string>();
//                        foreach (ErrorRecord err in this.PowerShell.Streams.Error)
//                        {
//                            errors.Add(err.ToString());
//                        }
//                        return new PSExecutionResult()
//                        {
//                            ComputerName = this.ComputerName,
//                            HadErrors = true,
//                            Results = null,
//                            Errors = errors
//                        };
//                    }
//                    else
//                    {
//                        return new PSExecutionResult()
//                        {
//                            ComputerName = this.ComputerName,
//                            HadErrors = false,
//                            Results = result,
//                            Errors = null
//                        };
//                    }
//                }
//                catch (Exception e)
//                {
//                    return new PSExecutionResult()
//                    {
//                        ComputerName = this.ComputerName,
//                        HadErrors = true,
//                        Errors = new List<string>() { e.ToString() },
//                        Results = null
//                    };
//                }
//                finally
//                {
//                    result.Clear();
//                }
//            }
//        }
//        #endregion
//    }
//    public class PSExecutionResult
//    {
//        #region Public properties
//        public string ComputerName { get; set; }
//        public object Results { get; set; }
//        public bool HadErrors { get; set; }
//        public List<string> Errors { get; set; }
//        #endregion
//    }

//    public class PowerShell
//    {
//        #region Public properties
//        public List<string> ComputerNames { get; }
//        public string Code { get; }
//        public CodeType CodeType { get; }
//        public Dictionary<string, object> Parameters { get; }
//        public List<string> ModulesTobeImported { get; }
//        public byte TimeoutInSeconds { get; }
//        public byte Throttle { get; }
//        #endregion

//        #region Constructor
//        public PowerShell(
//            List<string> computernames,
//            string code,
//            CodeType codeType,
//            Dictionary<string, object> parameters,
//            List<string> modulesTobeImported,
//            byte timeoutInSeconds = 30,
//            byte throttle = 4 //Convert.ToByte(Environment.ProcessorCount);
//        )
//        {
//            this.ComputerNames = computernames;
//            this.Code = code;
//            this.CodeType = codeType;
//            this.Parameters = parameters;
//            this.ModulesTobeImported = modulesTobeImported;
//            this.TimeoutInSeconds = timeoutInSeconds;
//            this.Throttle = throttle;
//        }
//        #endregion

//        #region Public methods
//        public ObservableCollection<PSExecutionResult> BeginInvoke()
//        {
//            Random lockObject = new Random();
//            ParallelOptions po = new ParallelOptions();
//            ObservableCollection<PSExecutionResult> results = new ObservableCollection<PSExecutionResult>();
//            po.MaxDegreeOfParallelism = this.Throttle;
//            var psCodeObjects = CreatePSCodeObjectsList();
//            Parallel.ForEach<PSCode>(psCodeObjects, po, psCode =>
//            {
//                var res = psCode.Invoke(this.Code, this.Parameters, this.ModulesTobeImported, this.TimeoutInSeconds);
//                lock (lockObject)
//                {
//                    results.Add(res);
//                }
//            });
//            return results;
//        }
//        #endregion

//        #region Private methods
//        private IEnumerable<PSCode> CreatePSCodeObjectsList()
//        {
//            foreach (var computer in ComputerNames)
//            {
//                yield return (new PSCode(computer));
//            }
//        }
//        #endregion
//    }

//    public class program
//    {
//        public static void Main()
//        {
//            program p = new program();
//            var a = p.call();
//        }
//        private dynamic call()
//        {
//            var ps = new PowerSharp.PowerShell(new List<string>() { "mslaptop" }, "get-service", CodeType.Script, null, null);
//            return ps.BeginInvoke();
//        }
//    }
//}