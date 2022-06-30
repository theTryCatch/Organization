//using Organization.PowerSharp;
//using System;
//using System.Collections.Generic;
//using System.Globalization;
//using System.IO;
//using System.Management.Automation.Host;
//using System.Management.Automation.Runspaces;
//using MSFT = System.Management.Automation;

//namespace OrgUtils.PowerSharp
//{
//    public class UsingRSPool
//    {
//        #region Public properties
//        public IEnumerable<string> ComputerNames { get; }
//        public string Code { get; }
//        public CodeType CodeType { get; }
//        public Dictionary<string, object> Parameters { get; }
//        public IEnumerable<FileInfo> ModulesTobeImported { get; }
//        public uint TimeoutInSeconds { get; }
//        public uint Throttle { get; }
//        #endregion

//        #region Constructor
//        public UsingRSPool(IEnumerable<string> computernames, string code, CodeType codeType, Dictionary<string, object> parameters, IEnumerable<FileInfo> modulesTobeImported, uint timeoutInSeconds = 30, uint throttle = 4)
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

//        public IEnumerable<PSExecutionResult> Invoke()
//        {
//            foreach (var computer in ComputerNames)
//            {
//                var iss = InitialSessionState.CreateDefault();
//                iss.ExecutionPolicy = Microsoft.PowerShell.ExecutionPolicy.Unrestricted;
//                iss.DisableFormatUpdates = true;
//                using (var runspacePool = RunspaceFactory.CreateRunspacePool(1, Environment.ProcessorCount, iss, new PSHost(new PSHostUserInterface(this))))
//                {
//                    using (var ps = MSFT.PowerShell.Create(iss))
//                    {
//                        ps.RunspacePool = runspacePool;
//                    }
//                }
//            }
//        }
//    }
//    internal class MyHost : PSHost, IHostSupportsInteractiveSession
//    {
//        public MyHost(UsingRSPool program)
//        {
//            this.program = program;
//        }

//        /// <summary>
//        /// A reference to the PSHost implementation.
//        /// </summary>
//        private UsingRSPool program;

//        /// <summary>
//        /// The culture information of the thread that created
//        /// this object.
//        /// </summary>
//        private CultureInfo originalCultureInfo =
//            System.Threading.Thread.CurrentThread.CurrentCulture;

//        /// <summary>
//        /// The UI culture information of the thread that created
//        /// this object.
//        /// </summary>
//        private CultureInfo originalUICultureInfo =
//            System.Threading.Thread.CurrentThread.CurrentUICulture;

//        /// <summary>
//        /// The identifier of this PSHost implementation.
//        /// </summary>
//        private static Guid instanceId = Guid.NewGuid();

//        /// <summary>
//        /// A reference to the implementation of the PSHostUserInterface
//        /// class for this application.
//        /// </summary>
//        private MyHostUserInterface myHostUserInterface = new MyHostUserInterface();

//        /// <summary>
//        /// A reference to the runspace used to start an interactive session.
//        /// </summary>
//        public Runspace pushedRunspace = null;

//        /// <summary>
//        /// Gets the culture information to use. This implementation
//        /// returns a snapshot of the culture information of the thread
//        /// that created this object.
//        /// </summary>
//        public override CultureInfo CurrentCulture
//        {
//            get { return this.originalCultureInfo; }
//        }

//        /// <summary>
//        /// Gets the UI culture information to use. This implementation
//        /// returns a snapshot of the UI culture information of the thread
//        /// that created this object.
//        /// </summary>
//        public override CultureInfo CurrentUICulture
//        {
//            get { return this.originalUICultureInfo; }
//        }

//        /// <summary>
//        /// Gets an identifier for this host. This implementation always
//        /// returns the GUID allocated at instantiation time.
//        /// </summary>
//        public override Guid InstanceId
//        {
//            get { return instanceId; }
//        }

//        /// <summary>
//        /// Gets a string that contains the name of this host implementation.
//        /// Keep in mind that this string may be used by script writers to
//        /// identify when your host is being used.
//        /// </summary>
//        public override string Name
//        {
//            get { return "MySampleConsoleHostImplementation"; }
//        }

//        /// <summary>
//        /// Gets an instance of the implementation of the PSHostUserInterface
//        /// class for this application. This instance is allocated once at startup time
//        /// and returned every time thereafter.
//        /// </summary>
//        public override PSHostUserInterface UI
//        {
//            get { return this.myHostUserInterface; }
//        }

//        /// <summary>
//        /// Gets the version object for this application. Typically this
//        /// should match the version resource in the application.
//        /// </summary>
//        public override Version Version
//        {
//            get { return new Version(1, 0, 0, 0); }
//        }

//        #region IHostSupportsInteractiveSession Properties

//        /// <summary>
//        /// Gets a value indicating whether a request
//        /// to open a PSSession has been made.
//        /// </summary>
//        public bool IsRunspacePushed
//        {
//            get { return this.pushedRunspace != null; }
//        }

//        /// <summary>
//        /// Gets or sets the runspace used by the PSSession.
//        /// </summary>
//        public Runspace Runspace
//        {
//            get { return this.program.myRunSpace; }
//            internal set { this.program.myRunSpace = value; }
//        }
//        #endregion IHostSupportsInteractiveSession Properties

//        /// <summary>
//        /// This API Instructs the host to interrupt the currently running
//        /// pipeline and start a new nested input loop. In this example this
//        /// functionality is not needed so the method throws a
//        /// NotImplementedException exception.
//        /// </summary>
//        public override void EnterNestedPrompt()
//        {
//            throw new NotImplementedException(
//                  "The method or operation is not implemented.");
//        }

//        /// <summary>
//        /// This API instructs the host to exit the currently running input loop.
//        /// In this example this functionality is not needed so the method
//        /// throws a NotImplementedException exception.
//        /// </summary>
//        public override void ExitNestedPrompt()
//        {
//            throw new NotImplementedException(
//                  "The method or operation is not implemented.");
//        }

//        /// <summary>
//        /// This API is called before an external application process is
//        /// started. Typically it is used to save state so that the parent
//        /// can restore state that has been modified by a child process (after
//        /// the child exits). In this example this functionality is not
//        /// needed so the method returns nothing.
//        /// </summary>
//        public override void NotifyBeginApplication()
//        {
//            return;
//        }

//        /// <summary>
//        /// This API is called after an external application process finishes.
//        /// Typically it is used to restore state that a child process has
//        /// altered. In this example, this functionality is not needed so
//        /// the method returns nothing.
//        /// </summary>
//        public override void NotifyEndApplication()
//        {
//            return;
//        }

//        /// <summary>
//        /// Indicate to the host application that exit has
//        /// been requested. Pass the exit code that the host
//        /// application should use when exiting the process.
//        /// </summary>
//        /// <param name="exitCode">The exit code that the
//        /// host application should use.</param>
//        public override void SetShouldExit(int exitCode)
//        {
//            this.program.ShouldExit = true;
//            this.program.ExitCode = exitCode;
//        }

//        #region IHostSupportsInteractiveSession Methods

//        /// <summary>
//        /// Requests to close a PSSession.
//        /// </summary>
//        public void PopRunspace()
//        {
//            Runspace = this.pushedRunspace;
//            this.pushedRunspace = null;
//        }

//        /// <summary>
//        /// Requests to open a PSSession.
//        /// </summary>
//        /// <param name="runspace">Runspace to use.</param>
//        public void PushRunspace(Runspace runspace)
//        {
//            this.pushedRunspace = Runspace;
//            Runspace = runspace;
//        }

//        #endregion IHostSupportsInteractiveSession Methods
//    }
//}
