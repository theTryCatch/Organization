using System.Management.Automation.Runspaces;
using MSFT = System.Management.Automation;
namespace Organization.PowerSharp
{
    public class PSCode
    {
        #region Properties
        /// <summary>
        /// PSCode enables you to run any PowerShell command, script on a local or remote machine with timeout feature.
        /// </summary>
        /// <remarks>
        /// This class implements PowerShell runspaces for execution and TPL (Task) for timeout feature.
        /// </remarks>
        public string ComputerName { get; }
        public string Code { get; }
        public CodeType CodeType { get; }
        public Dictionary<string, object> Parameters { get; }
        public List<FileInfo> ModulesTobeImported { get; }
        public uint TimeoutInSeconds { get; }



        private CancellationTokenSource? _cts { get; set; }
        private Task<PSExecutionResult>? _task { get; set; }
        private MSFT.PowerShell? _powershell { get; set; }
        private Runspace? _runspace { get; set; }

        #endregion

        #region Constructor
        /// <summary>
        /// Construct a new <c>PSCode</c>instance
        /// </summary>
        /// <param name="computername">Target computer name</param>
        /// <param name="code">PowerShell code (cmdlet, script, script file path)</param>
        /// <param name="codeType">What type of 'code' you provided?</param>
        /// <param name="parameters">specify list of parameters for the PowerShell 'code'</param>
        /// <param name="modulesTobeImported">Provide list of modules you wanted to import into the PowerShell environment.</param>
        /// <param name="timeoutInSeconds">Specify the timeout in seconds. Defaults to 30 seconds</param>
        public PSCode(
            string computername,
            string code, CodeType codeType,
            Dictionary<string, object> parameters,
            List<FileInfo> modulesTobeImported,
            uint timeoutInSeconds = 30
        )
        {
            this.ComputerName = computername;
            this.Code = code;
            this.CodeType = codeType;
            this.Parameters = parameters;
            this.ModulesTobeImported = modulesTobeImported;
            this.TimeoutInSeconds = timeoutInSeconds;



            this._cts = new CancellationTokenSource();
            this._cts.CancelAfter(TimeSpan.FromSeconds(this.TimeoutInSeconds));
            this._powershell = MSFT.PowerShell.Create();
            if (computername.ToUpper() == Environment.MachineName.ToUpper())
            {
                this._runspace = RunspaceFactory.CreateRunspace();
            }
            else
            {
                this._runspace = RunspaceFactory.CreateRunspace(new WSManConnectionInfo(new Uri($"http://{this.ComputerName}:{5985}/WSMAN")));
            }
            this._powershell.Runspace = this._runspace;


            if (modulesTobeImported != null)
            {
                foreach (var item in modulesTobeImported)
                {
                    this._powershell.AddCommand("Import-Module").AddArgument(item.ToString());
                }
            }
            if (codeType == CodeType.Cmdlet)
                this._powershell.AddCommand(code);
            else if (codeType == CodeType.Script)
                this._powershell.AddScript(code);
            else
                this._powershell.AddScript(File.ReadAllText(code));

            if (parameters != null)
                this._powershell.AddParameters(parameters);
        }
        #endregion

        #region Public methods
        /// <summary>
        /// This method invokes PowerShell synchronously and time out the execution based on the provided <see cref="TimeoutInSeconds"/> parameter.
        /// </summary>
        /// <remarks>This method is never expected to throw any errors</remarks>
        public PSExecutionResult Invoke()
        {
            PSExecutionResult? _executionResults;
            List<string> _errors = new List<string>();

            Func<PSExecutionResult> func = new Func<PSExecutionResult>(IgnitePowerShellInvocation);
            using (this._task)
            {
                using (this._cts)
                {
                    try
                    {
                        this._task = Task.Run(func, this._cts.Token);
                        this._task.Wait(this._cts.Token);
                        _executionResults = this._task.Result;
                    }
                    catch (OperationCanceledException)
                    {
                        _errors.Add($"Execution timeout with in {this.TimeoutInSeconds} second(s)");
                        _executionResults = null;
                        return new PSExecutionResult()
                        {
                            ComputerName = this.ComputerName,
                            Errors = _errors,
                            HadErrors = true,
                            Results = null
                        };
                    }
                    catch (Exception e)
                    {
                        _errors.Add(e.ToString());
                        _executionResults = null;
                        return new PSExecutionResult()
                        {
                            ComputerName = this.ComputerName,
                            Errors = _errors,
                            HadErrors = true,
                            Results = null
                        };
                    }
                }
                this._cts = null;
            }
            this._task = null;
            return _executionResults;
        }
        #endregion

        #region Private methods
        private PSExecutionResult IgnitePowerShellInvocation()
        {
            PSExecutionResult _executionResults;
            List<string> _errors = new List<string>();
            try
            {
                using (this._powershell)
                {
                    using (this._runspace)
                    {
                        this._runspace.Open();
                        var result = this._powershell.Invoke();

                        //If the executed code is just a command then the PowerShell.Invoke() method throws an exception but
                        //if the code is a script and contains any errors then it will not indicate anyway about the exceptions.
                        //Hence we the logic is around HadErrors property data.
                        if (this._powershell.HadErrors)
                        {
                            foreach (var item in (from error in this._powershell.Streams.Error select error.Exception.Message))
                            {
                                _errors.Add(item);
                            }
                            _executionResults = new PSExecutionResult()
                            {
                                ComputerName = this.ComputerName,
                                HadErrors = true,
                                Results = null,
                                Errors = _errors
                            };
                        }
                        else
                        {
                            _executionResults = new PSExecutionResult()
                            {
                                ComputerName = this.ComputerName,
                                HadErrors = false,
                                Results = result,
                                Errors = null
                            };
                        }
                    }
                    this._runspace = null;
                }
                this._powershell = null;
            }
            catch (Exception e)
            {
                _errors.Add(e.ToString());
                _executionResults = new PSExecutionResult()
                {
                    ComputerName = this.ComputerName,
                    HadErrors = true,
                    Results = null,
                    Errors = _errors
                };
            }
            return _executionResults;
        }
        #endregion
    }
}