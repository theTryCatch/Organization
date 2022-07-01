//using Organization.PowerSharp;
//using System;
//using System.Collections.Generic;
//using System.Collections.ObjectModel;
//using System.Globalization;
//using System.IO;
//using System.Management.Automation;
//using System.Management.Automation.Host;
//using System.Management.Automation.Runspaces;
//using System.Threading;
//using System.Threading.Tasks;
//using MSFT = System.Management.Automation;

//namespace OrgUtils.PowerSharp
//{
//    public class UsingRSPool
//    {
//        #region Public properties
//        public List<string> ComputerNames { get; }
//        public string Code { get; }
//        public CodeType CodeType { get; }
//        public Dictionary<string, object> Parameters { get; }
//        public List<FileInfo> ModulesTobeImported { get; }
//        public uint TimeoutInSeconds { get; }
//        public uint Throttle { get; }
//        #endregion

//        #region Constructor
//        public UsingRSPool(List<string> computernames, string code, CodeType codeType, Dictionary<string, object> parameters, List<FileInfo> modulesTobeImported, uint timeoutInSeconds = 30, uint throttle = 4)
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
//        ~UsingRSPool()
//        {
//            GC.Collect();
//        }

//        public void Invoke()
//        {
//            //List<PSDataCollection<PSObject>> cps = new List<PSDataCollection<PSObject>>();
            
//            foreach (var computer in ComputerNames)
//            {
//                using (var runspacePool = RunspaceFactory.CreateRunspacePool(1, 4))
//                {
//                    //runspacePool.InitialSessionState.ExecutionPolicy = Microsoft.PowerShell.ExecutionPolicy.Unrestricted;
//                    runspacePool.ThreadOptions = PSThreadOptions.ReuseThread;
//                    var ps = MSFT.PowerShell.Create();
                    
//                        //using (CancellationTokenSource cts = new CancellationTokenSource())
//                        //{
//                            //cts.CancelAfter(TimeSpan.FromSeconds(2));

//                    ps.RunspacePool = runspacePool;
//                    ps.AddScript($"invoke-command -computer {computer} -scriptBlock {{get-service}}");
//                    ps.RunspacePool.Open();
//                    var endTime = DateTime.Now.AddSeconds(1).ToUniversalTime();
//                    var handler = ps.BeginInvoke();
//                    while (DateTime.Now.ToUniversalTime() <= endTime)
//                    {
//                        if (handler.IsCompleted)
//                            break;
//                    }
//                    if (handler.IsCompleted)
//                        ps.EndInvoke(handler);
//                    else
//                        Console.WriteLine("Timedout");
//                    Task<Collection<PSObject>> task = Task.Run(() => ps.Invoke(), cts.Token);
//                    //using (task)
//                    //{
//                    //    task.Wait(cts.Token);
//                    //    cps.Add(task.Result);
//                    //}

//                    //task = null;
//                    //}
//                }
//            }
//            //return cps;
//        }
//    }
//    public class MyClass
//    {
//        public static void Main()
//        {
//            var a = new UsingRSPool(new List<string>() { "mslaptop" }, "get-service", CodeType.Cmdlet, null, null);
//            a.Invoke();
//        }
//    }
//}
