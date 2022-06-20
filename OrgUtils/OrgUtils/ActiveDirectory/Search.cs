using Organization.LDAP;
using System.DirectoryServices;

namespace Organization.ActiveDirectory
{
    public static class Search
    {
        public static dynamic FindADObject(Filter filter, bool findAll = false)
        {
            DirectorySearcher searcher = GetRootDSESearcherObject();

            searcher.Filter = filter.LDAPFilter;
            if (findAll == false)
                return searcher.FindOne();
            else
                return searcher.FindAll();
        }

        public static DirectorySearcher GetRootDSESearcherObject()
        {
            return new DirectorySearcher((new DirectoryEntry("GC://" + ((new DirectoryEntry("LDAP://RootDSE")).Properties["rootDomainNamingContext"].Value.ToString()))));
        }
    }
}