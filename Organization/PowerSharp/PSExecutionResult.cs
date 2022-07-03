using System.Collections.ObjectModel;
using System.Management.Automation;

namespace Organization.PowerSharp
{
    public class PSExecutionResult
    {
        public string? ComputerName { get; set; }
        public Collection<PSObject>? Results { get; set; }
        public bool HadErrors { get; set; }
        public IEnumerable<string>? Errors { get; set; }
    }
}
