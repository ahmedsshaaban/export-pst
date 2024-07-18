using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ExportPst_multiPages
{
    /// <summary>
    /// Interaction logic for login.xaml
    /// </summary>
    public partial class login : Page
    {
        public login()
        {
            InitializeComponent();

            


        }
       

        private void loginb_Click(object sender, RoutedEventArgs e)
        {
            (App.Current as App).username = username.Text;

            (App.Current as App).password = password.Password;
            MainWindow Form = Application.Current.Windows[0] as MainWindow;
            if (MailBoxSearch.login("ahmedsh@extremetech.tk", "Hash98$@@$12"))
            {
                Form.MainW.Content = new Mailboxes();
            }
            else
            {
                error.Visibility = Visibility.Visible;
            }
            
            
        }
    }
}
