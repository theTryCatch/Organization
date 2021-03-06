namespace Organization.LDAP
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