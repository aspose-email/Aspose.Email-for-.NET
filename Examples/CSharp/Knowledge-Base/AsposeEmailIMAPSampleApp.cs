using Aspose.Email;
using Aspose.Email.Clients;
using Aspose.Email.Clients.Imap;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Knowledge.Base
{
    public partial class AsposeEmailIMAPSampleApp : Form
    {
        public AsposeEmailIMAPSampleApp()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // ExStart:AsposeEmailIMAP
            // Creates an instance of the class ImapClient by specified the host, username and password
            ImapClient client = new ImapClient("localhost", "username", "password");

            try
            {
                client.SelectFolder(ImapFolderInfo.InBox);
                String strTemp;
                strTemp = "You have " + client.CurrentFolder.TotalMessageCount.ToString() + " messages in your account.";
                // Gets number of messages in the folder, Disconnects to imap server.
                MessageBox.Show(strTemp);
            }
            catch (ImapException ex)
            {
                MessageBox.Show(ex.ToString());
            }
            // ExEnd:AsposeEmailIMAP
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // Creates an instance of the class ImapClient by specified the host, username and password
            ImapClient client = new ImapClient("localhost", "username", "password");

            // ExStart:SSLEnabledServer
            // Set security mode
            client.SecurityOptions = SecurityOptions.SSLImplicit;
            // ExEnd:SSLEnabledServer

            try
            {
                client.SelectFolder(ImapFolderInfo.InBox);
                String strTemp;
                strTemp = "You have " + client.CurrentFolder.TotalMessageCount.ToString() + " messages in your account.";
                // Gets number of messages in the folder, Disconnects to imap server.
                MessageBox.Show(strTemp);
            }
            catch (ImapException ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
    }
}
