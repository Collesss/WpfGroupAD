using System.DirectoryServices;

namespace WpfAppAddUsersToGroups.Extensions
{
    public static class SearchResultExtensions
    {
        public static string GetProp(this SearchResult searchResult, string prop) =>
            searchResult.Properties[prop][0].ToString();
    }
}
