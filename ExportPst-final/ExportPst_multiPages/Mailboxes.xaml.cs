using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
    /// Interaction logic for Mailboxes.xaml
    /// </summary>
    public partial class Mailboxes : Page
    {
        public ListView SelectUsers { get; set; }
        public Mailboxes()
        {
            InitializeComponent();

            if (selectUsers.Items.Count == 0)
                txtFilter.IsEnabled = false;

        }
        private bool UserFilter(object item)
        {
            if (String.IsNullOrEmpty(txtFilter.Text))
                return true;
            else {
                if((item as Mailbox).Name.IndexOf(txtFilter.Text, StringComparison.OrdinalIgnoreCase) >= 0)
                return true;
                if ((item as Mailbox).Database.IndexOf(txtFilter.Text, StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
                if ((item as Mailbox).EmailAddress.IndexOf(txtFilter.Text, StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
            }

            return false;
                
        }

        private void txtFilter_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            if (selectUsers.ItemsSource != null)
                CollectionViewSource.GetDefaultView(selectUsers.ItemsSource).Refresh();
        }

        private async void getSource_Click(object sender, RoutedEventArgs e)
        {


           var ser = new MailBoxSearch("ahmedsh@extremetech.tk", "Hash98$@@$12");
            //var allMBXs = await ser.getall();
            //selectUsers.ItemsSource = allMBXs;
            //CollectionView view = (CollectionView)CollectionViewSource.GetDefaultView(selectUsers.ItemsSource);
            //view.Filter = UserFilter;
            //(App.Current as App).SelectUsers = selectUsers;


            if (selectsource.Text == "All")
            {

                txtFilter.Text = "";
                source.ItemsSource = null;
                source.Visibility = Visibility.Hidden;
             //   var watch = System.Diagnostics.Stopwatch.StartNew();
                var allMBXs = await ser.getall();
                selectUsers.ItemsSource = allMBXs;
              //  watch.Stop();
               // var elapsedMs = watch.ElapsedMilliseconds;
              //  MessageBox.Show(elapsedMs.ToString());
            }

            if (selectsource.Text == "Group")
            {
                txtFilter.Text = "";
                source.Visibility = Visibility.Visible;
                selectUsers.ItemsSource = null;
                var groups = await ser.getGroups();
                source.ItemsSource = groups;
            }
            if (selectsource.Text == "Database")
            {
                txtFilter.Text = "";
                source.Visibility = Visibility.Visible;
                selectUsers.ItemsSource = null;

                var Databases = await ser.getDatabase();
                source.ItemsSource = Databases;
            }
            if (selectsource.Text == "OU")
            {
                txtFilter.Text = "";
                source.Visibility = Visibility.Visible;
                selectUsers.ItemsSource = null;

                var OUs = await ser.getOUs();
                source.ItemsSource = OUs;
            }
            if (selectUsers.ItemsSource != null)
            {
                txtFilter.IsEnabled = true;
                CollectionView view = (CollectionView)CollectionViewSource.GetDefaultView(selectUsers.ItemsSource);
                view.Filter = UserFilter;
            }
            
               
            ser.RunSpace.Close();
            ser.RunSpace.Dispose();
            ser.PowerShell.Dispose();
           
        }

        private async void getUsers_Click(object sender, RoutedEventArgs e)
        {
            var search = new MailBoxSearch("ahmedsh@extremetech.tk", "Hash98$@@$12");
            if (selectsource.Text == "Database") {
                List<string> list = new List<string>() ;
                foreach (var selectedItem in source.SelectedItems)
                {
                    list.Add(selectedItem.ToString());
                }

                selectUsers.ItemsSource= await search.getallx(list, "DB");
            }

            if (selectsource.Text == "OU")
            {
                List<string> list = new List<string>();
                foreach (var selectedItem in source.SelectedItems)
                {
                    list.Add(selectedItem.ToString());
                }

                selectUsers.ItemsSource = await search.getallx(list,"OU");
            }
            if (selectsource.Text == "Group")
            {
                List<string> list = new List<string>();
                foreach (var selectedItem in source.SelectedItems)
                {
                    list.Add(selectedItem.ToString());
                }

                selectUsers.ItemsSource = await search.getallx(list, "Group");
            }
            if (selectUsers.ItemsSource != null)
            {
                txtFilter.IsEnabled = true;
                CollectionView view = (CollectionView)CollectionViewSource.GetDefaultView(selectUsers.ItemsSource);
                view.Filter = UserFilter;
            }
            search.RunSpace.Close();
            search.RunSpace.Dispose();
            search.PowerShell.Dispose();
        }
    }
}
