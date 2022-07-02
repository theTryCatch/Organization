using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Threading;
using System.Threading.Tasks;
using MSFT = System.Management.Automation;
namespace OrgUtils.PowerSharp
{
    public sealed class PSCode : IDisposable
    {
        #region Properties
        public string ComputerName { get; }
        public string Code { get; }
        public CodeType CodeType { get; }
        public Dictionary<string, object> Parameters { get; }
        public List<FileInfo> ModulesTobeImported { get; }
        public uint TimeoutInSeconds { get; }



        private CancellationTokenSource _cts { get; set; }
        private Task<PSExecutionResult> _task { get; set; }
        private MSFT.PowerShell _powershell { get; set; }
        private Runspace _runspace { get; set; }
        private List<string> _errors { get; set; } = new List<string>();
        private Collection<PSObject> _results { get; set; }

        #endregion

        #region Constructor
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
        public PSExecutionResult Invoke()
        {
            PSExecutionResult result = null;
            Func<PSExecutionResult> func = new Func<PSExecutionResult>(IgnitePowerShellInvocation);
            using (this._task)
            {
                using (this._cts)
                {
                    try
                    {
                        this._task = Task.Run(func, this._cts.Token);
                        this._task.Wait(this._cts.Token);
                        result = this._task.Result;
                    }
                    catch (OperationCanceledException)
                    {
                        this._errors.Add($"Execution timeout with in {this.TimeoutInSeconds} second(s)");
                        this._results = null;
                        return new PSExecutionResult()
                        {
                            ComputerName = this.ComputerName,
                            Errors = this._errors,
                            HadErrors = true,
                            Results = null
                        };
                    }
                    catch (Exception e)
                    {
                        this._errors.Add(e.ToString());
                        this._results = null;
                        return new PSExecutionResult()
                        {
                            ComputerName = this.ComputerName,
                            Errors = this._errors,
                            HadErrors = true,
                            Results = null
                        };
                    }
                }
                this._cts = null;
            }
            this._task = null;
            return result;
        }
        #endregion

        #region Private methods
        private PSExecutionResult IgnitePowerShellInvocation()
        {
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
                            foreach (var item in (from error in this._powershell.Streams.Error select error.ToString()))
                            {
                                this._errors.Add(item);
                            }
                        }
                        else
                        {
                            this._results = result;
                        }
                    }
                    this._runspace = null;
                }
                this._powershell = null;
            }
            catch (Exception e)
            {
                this._errors.Add(e.ToString());
                this._results = null;
            }
            if (this._errors.Count > 0)
            {
                return new PSExecutionResult()
                {
                    ComputerName = this.ComputerName,
                    HadErrors = true,
                    Results = null,
                    Errors = this._errors
                };
            }
            else
            {
                return new PSExecutionResult()
                {
                    ComputerName = this.ComputerName,
                    HadErrors = false,
                    Results = this._results,
                    Errors = null
                };
            }
        }
        #endregion

        #region Interface implementation
        public void Dispose()
        {
            if(this._errors != null)
                this._errors.Clear();

            if (this._results != null)
                this._results.Clear();
        }
        #endregion
    }
}