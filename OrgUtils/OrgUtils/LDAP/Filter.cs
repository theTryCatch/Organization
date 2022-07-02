using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OrgUtls.LDAP
{
    public class Filter
    {
        public string LDAPFilter { get; }
        public Filter(string identity, List<string> propertyName)
        {
            string filterSegment = $"(&(objectClass=*)(|";
            foreach (var item in propertyName)
            {
                filterSegment += $"{item}={identity}";
            }
            filterSegment += "))";

            this.LDAPFilter = filterSegment;
        }
    }
}