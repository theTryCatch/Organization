using Organization.PowerSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OrgUtils.PowerSharp
{
    public class MyClass
    {
        public static void Main()
        {
            var a = new PowerSharp.PowerShell(new List<string>() { "asdfsdf" }, "get-service; sleep 10", CodeType.Script, null, null);
            var b = a.BeginInvoke();
            Console.Read();
        }
    }
}
