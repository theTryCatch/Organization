using System.Collections.Generic;
using System.Management.Automation;

namespace OrgUtils.PowerSharp
{
    public class PSExecutionResult
    {
        public string ComputerName { get; set; }
        public object Results { get; set; }
        public bool HadErrors { get; set; }
        public IEnumerable<string> Errors { get; set; }
    }
}
