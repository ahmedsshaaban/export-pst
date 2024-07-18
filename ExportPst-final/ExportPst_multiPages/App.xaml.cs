using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace ExportPst_multiPages
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        public string username { get; set; }
        public string password { get; set; }
        public List<Mailbox> mailbx { get; set; }
        public ListView SelectUsers { get; set; }
        
        public TextBox TxtFilter { get; set; }

    }
}
