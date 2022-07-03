using System.Collections.ObjectModel;
using System.Collections.Specialized;

namespace Organization.PowerSharp
{
    public class PowerShell
    {
        #region Public properties
        public List<string> ComputerNames { get; }
        public string Code { get; }
        public CodeType CodeType { get; }
        public Dictionary<string, object> Parameters { get; }
        public List<FileInfo> ModulesTobeImported { get; }
        public byte TimeoutInSeconds { get; }
        public byte Throttle { get; }
        #endregion

        #region Constructor
        public PowerShell(
            List<string> computernames,
            string code,
            CodeType codeType,
            Dictionary<string, object> parameters,
            List<FileInfo> modulesTobeImported,
            byte timeoutInSeconds = 30,
            byte throttle = 0
        )
        {
            if (throttle == 0)
                throttle = Convert.ToByte(Environment.ProcessorCount);

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
            var lockObject = this;
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
        public ObservableCollection<PSExecutionResult> BeginInvoke(NotifyCollectionChangedEventHandler ExecutionCompleted)
        {
            var lockObject = this;
            ParallelOptions po = new ParallelOptions();
            ObservableCollection<PSExecutionResult> results = new ObservableCollection<PSExecutionResult>();
            results.CollectionChanged += ExecutionCompleted;
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
}
