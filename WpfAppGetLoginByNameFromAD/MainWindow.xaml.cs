using System.DirectoryServices;
using System.Linq;
using System.Windows;
using WpfAppGetLoginByNameFromAD.Extensions;

namespace WpfAppGetLoginByNameFromAD
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void ButtonFind_Click(object sender, RoutedEventArgs e)
        {
            DirectoryEntry directoryEntry = new DirectoryEntry("LDAP://DC=office,DC=crocusgroup,DC=ru");

            //TextBoxNamesInput.IsEnabled = false;

            string[] names = TextBoxNamesInput.Text.Split("\r\n")
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .ToArray();

            TextBoxNamesInput.Text = string.Join("\r\n", names);

            (string login, string mail, string displayName)[] findLoginsAndEmails = names.Select(name =>
            {
                DirectorySearcher dirSearch = new DirectorySearcher(directoryEntry, $"(&(objectClass=user)(objectCategory=person)(anr={name}))");
                SearchResultCollection searchResColl = dirSearch.FindAll();

                if (searchResColl.Count == 0)
                    return ("Not find.", "Not find.", "Not find.");
                else if (searchResColl.Count > 1)
                    return ("Find more 1 result.", "Find more 1 result.", "Find more 1 result.");

                SearchResult searchResult = searchResColl[0];

                return (searchResult.GetProp("sAMAccountName"), searchResult.GetProp("mail"), searchResult.GetProp("displayName"));
            }).ToArray();

            TextBoxOutputLogins.Text = string.Join('\n', findLoginsAndEmails.Select(lae => lae.login));
            TextBoxOutputDisplayNames.Text = string.Join("\n", findLoginsAndEmails.Select(lae => lae.displayName));
            TextBoxOutputMails.Text = string.Join('\n', findLoginsAndEmails.Select(lae => lae.mail));
            TextBoxOutputMailsFormattedOutlook.Text = string.Join("; ", findLoginsAndEmails.Select(lae => lae.mail));
            TextBoxOutputMailsFormattedHelp.Text = string.Join(",", findLoginsAndEmails.Select(lae => lae.mail));

            //TextBoxNamesInput.IsEnabled = true;
        }

        private void ButtonAddToGroup_Click(object sender, RoutedEventArgs e)
        {
            string group = TextBoxGroupInput.Text;

            DirectoryEntry directoryEntry = new DirectoryEntry("LDAP://DC=office,DC=crocusgroup,DC=ru");
            DirectorySearcher dirSearchGroup = new DirectorySearcher(directoryEntry, $"(&(objectClass=group)(objectCategory=group)(cn={group}))");

            SearchResultCollection searchResCollGroup = dirSearchGroup.FindAll();

            if (searchResCollGroup.Count > 2)
            {
                MessageBox.Show("Find more 1 group with this name.");
                return;
            }
            else if (searchResCollGroup.Count == 0)
            {
                MessageBox.Show("Not find group with this name.");
                return;
            }

            string dnGroup = searchResCollGroup[0].GetProp("distinguishedname");

            DirectorySearcher dirSearchMemberGroup = new DirectorySearcher(directoryEntry, $"(&(objectClass=user)(objectCategory=user)(memberOf={dnGroup}))");

            SearchResultCollection searchResCollUsersGroup = dirSearchMemberGroup.FindAll();

            string[] loginUsersGroup = new string[searchResCollUsersGroup.Count];

            for (int i = 0; i < searchResCollUsersGroup.Count; i++)
            {
                loginUsersGroup[i] = searchResCollUsersGroup[i].GetProp("sAMAccountName");
            }


            string[] names = TextBoxNamesInput.Text.Split("\r\n")
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .ToArray();

            DirectoryEntry ent = new DirectoryEntry($"LDAP://{dnGroup}");

            string[] resultAddToGroup = names.Select(name =>
            {
                DirectorySearcher dirSearch = new DirectorySearcher(directoryEntry, $"(&(objectClass=user)(objectCategory=person)(anr={name}))");
                SearchResultCollection searchResColl = dirSearch.FindAll();

                if (searchResColl.Count == 0)
                    return "Not find.";
                else if (searchResColl.Count > 1)
                    return "Find more 1 result.";

                SearchResult searchResult = searchResColl[0];
                string userLogin = searchResult.GetProp("sAMAccountName");
                string userdn = searchResult.GetProp("distinguishedname");

                if (loginUsersGroup.Contains(userLogin))
                    return $"{userLogin} already added.";

                ent.Properties["member"].Add(userdn);

                return $"{userLogin} ADD";
            }).ToArray();

            ent.CommitChanges();

            new AddToGroupWindow(group, string.Join("\r\n", names), string.Join("\r\n", resultAddToGroup)).ShowDialog();
        }
    }
}