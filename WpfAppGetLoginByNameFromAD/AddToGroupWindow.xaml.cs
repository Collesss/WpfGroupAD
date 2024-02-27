using System.Windows;

namespace WpfAppGetLoginByNameFromAD
{
    /// <summary>
    /// Логика взаимодействия для AddToGroupWindow.xaml
    /// </summary>
    public partial class AddToGroupWindow : Window
    {


        public AddToGroupWindow(string group, string users, string statusToAddGroup)
        {
            InitializeComponent();

            TextBoxOutputGroup.Text = group;
            TextBoxOutputUsers.Text = users;
            TextBoxOutputStatusAddToGroup.Text = statusToAddGroup;
        }
    }
}
