using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace ExportPst_multiPages
{
     class MailBoxSearch
    {
        public PowerShell PowerShell { get; private set; }

        public Runspace RunSpace { get; private set; }

        private List<Mailbox> MailboxList;
       // public List<string> Groups { get; private set; }
        public MailBoxSearch(string username, string passwordChars)
        {


            try
            {
                SecureString password = new SecureString();

                foreach (char x in passwordChars) { password.AppendChar(x); }
                PSCredential credential = new PSCredential(username, password);
                WSManConnectionInfo connectionInfo = new WSManConnectionInfo(new Uri("http://Exchange1.extremetech.tk/powershell"), "http://schemas.microsoft.com/powershell/Microsoft.Exchange", credential);
                this.RunSpace = RunspaceFactory.CreateRunspace(connectionInfo);
                this.PowerShell = PowerShell.Create();
                this.PowerShell.Runspace = this.RunSpace;
                this.RunSpace.Open();


            }
            catch (Exception e)
            {

            }

            }
        public static Boolean  login(string username , string passwordChars)
        {
           
            try
            {
                SecureString password = new SecureString();

                foreach (char x in passwordChars) { password.AppendChar(x); }
                PSCredential credential = new PSCredential(username, password);
                WSManConnectionInfo connectionInfo = new WSManConnectionInfo(new Uri("http://Exchange1.extremetech.tk/powershell"), "http://schemas.microsoft.com/powershell/Microsoft.Exchange", credential);
                Runspace runspace = RunspaceFactory.CreateRunspace(connectionInfo);

                runspace.Open();
                runspace.Close();

                return true;

            }
            catch(Exception e)
            {
                return false;
            }
        }

        public  async Task<List<Mailbox>> getall()
        {
            try
            {
                this.MailboxList = new List<Mailbox>();
                PSCommand command = new PSCommand();

                command.AddCommand("Get-mailbox");
                command.AddParameter("ResultSize", "unlimited");
                //using select-object instead of filtering all prop like:
                //      MailboxList.Add(new Mailbox() {Name=mailbox.Properties.ElementAt(199).Value.ToString(),EmailAddress= mailbox.Properties.ElementAt(183).Value.ToString(),Database= mailbox.Properties.ElementAt(0).Value.ToString() });
                // is X2 faster


                command.AddCommand("select-object");
                command.AddParameter("Property", new string[3] { "Name", "PrimarySmtpAddress", "Database"});
                this.PowerShell.Commands = command;
                var mailboxes= await Task.Run(() => PowerShell.Invoke());
                foreach (var mailbox in mailboxes)
                {
                    MailboxList.Add(new Mailbox() {Name=mailbox.Properties.ElementAt(0).Value.ToString(),EmailAddress= mailbox.Properties.ElementAt(1).Value.ToString(),Database= mailbox.Properties.ElementAt(2).Value.ToString() });

                }

                return MailboxList;
            } catch(Exception e)
            {
                throw new Exception();
            }

        }

       
        public async Task<List<Mailbox>> getallx(List<string> sources,string type)
        {
            try
            {
                this.MailboxList = new List<Mailbox>();
                foreach (var source in sources)
                {
                    PSCommand command = new PSCommand();

                    if(type != "Group")
                    {
                        command.AddCommand("Get-mailbox");
                        command.AddParameter("ResultSize", "unlimited");

                    }
                    if(type == "Group")
                    {
                        command.AddCommand("Get-DistributionGroupMember");
                        command.AddParameter("ResultSize", "unlimited");
                        command.AddParameter("Identity", source);

                    }
                    if (type == "OU")
                     command.AddParameter("OrganizationalUnit", source);
                    if (type == "DB")
                        command.AddParameter("Database", source);
                    command.AddCommand("select-object");
                    command.AddParameter("Property", new string[3] { "Name", "PrimarySmtpAddress", "Database" });
                    this.PowerShell.Commands = command;
                    var mailboxes = await Task.Run(() => PowerShell.Invoke());
                    foreach (var mailbox in mailboxes)
                    {
                        MailboxList.Add(new Mailbox() { Name = mailbox.Properties.ElementAt(0).Value.ToString(), EmailAddress = mailbox.Properties.ElementAt(1).Value.ToString(), Database = mailbox.Properties.ElementAt(2).Value.ToString() });

                    }

                }
                return MailboxList;

            }
            catch (Exception e)
            {
                throw new Exception();
            }

        }

        public async Task<List<string>> getGroups()
        {
            List<string> groupNames = new List<string>();
            PSCommand command = new PSCommand();
            command.AddCommand("Get-DistributionGroup");
            command.AddParameter("ResultSize", "unlimited");
            this.PowerShell.Commands = command;
            var groups = await Task.Run(() => PowerShell.Invoke());
            foreach (var group in groups)
            {
                groupNames.Add(group.Properties.ElementAt(1).Value.ToString());
            }
            return groupNames;
        }
        public async Task<List<string>> getDatabase()
        {
            List<string> Databases = new List<string>();
            PSCommand command = new PSCommand();
            command.AddCommand("Get-MailboxDatabase");
            
            this.PowerShell.Commands = command;
            var DBs = await Task.Run(() => PowerShell.Invoke());
            foreach (var DB in DBs)
            {
                Databases.Add(DB.Properties.ElementAt(106).Value.ToString());
            }
            return Databases;
        }
        public async Task<List<string>> getOUs()
        {
            List<string> OUs = new List<string>();
            PSCommand command = new PSCommand();
            command.AddCommand("Get-OrganizationalUnit");

            this.PowerShell.Commands = command;
            var OrgUnits = await Task.Run(() => PowerShell.Invoke());
            foreach (var OU in OrgUnits)
            {
                OUs.Add(OU.Properties.ElementAt(17).Value.ToString());
            }
            return OUs;
        }

       
        //public static ObservableCollection<String> getMailboxesList()
        //{
        //    //MainWindow Form = Application.Current.Windows[0] as MainWindow;
        //    SecureString password = new SecureString();
        //    //string str_password = Form.passwordc.Password;
        //    //string username = Form.usernamec.Text;
        //    string str_password = "Hash98$@@$12";
        //    string username = "ahmedsh@extremetech.tk";
        //    foreach (char x in str_password) { password.AppendChar(x); }
        //    PSCredential credential = new PSCredential(username, password);

        //    WSManConnectionInfo connectionInfo = new WSManConnectionInfo(new Uri("http://Exchange1.extremetech.tk/powershell"), "http://schemas.microsoft.com/powershell/Microsoft.Exchange", credential);

        //    //connectionInfo.AuthenticationMechanism = AuthenticationMechanism.Basic;

        //    Runspace runspace = RunspaceFactory.CreateRunspace(connectionInfo);


        //    PowerShell powershell = PowerShell.Create();

        //    PSCommand command = new PSCommand();

        //    command.AddCommand("Get-mailbox");
        //    command.AddParameter("ResultSize", "unlimited");
        //    command.AddCommand("select-object");
        //    command.AddParameter("Property", new string[3] { "Name", "PrimarySmtpAddress", "Database" });
        //    powershell.Commands = command;

        //    runspace.Open();
        //    powershell.Runspace = runspace;
        //    var aa = powershell.Invoke();
        //    ObservableCollection<string> mailBoxes = new ObservableCollection<string>();
        //    foreach (var mailbox in aa)
        //    {
        //        // mailBoxes.Add(new Mailbox() { Name = mailbox.Properties.ElementAt(0).Value.ToString(), EmailAddress = mailbox.Properties.ElementAt(1).Value.ToString(), Database = mailbox.Properties.ElementAt(2).Value.ToString() });
        //        mailBoxes.Add(mailbox.Properties.ElementAt(0).Value.ToString());
        //    }
        //    runspace.Close();
        //    return mailBoxes;

        //}
    }
}
