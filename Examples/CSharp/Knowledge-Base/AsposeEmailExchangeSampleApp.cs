using Aspose.Email.Clients.Exchange;
using Aspose.Email.Clients.Exchange.Dav;
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
    public partial class AsposeEmailExchangeSampleApp : Form
    {
        public AsposeEmailExchangeSampleApp()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                // ExStart:AsposeEmailExchange
                // Clear the items in the listbox
                lstMessages.Items.Clear();

                // Create instance of ExchangeClient class by giving credentials and Call ListMessages method to list messages info from Inbox
                ExchangeClient client = new ExchangeClient(txtExchangeServer.Text, txtUsername.Text, txtPassword.Text, txtDomain.Text);
                ExchangeMessageInfoCollection msgCollection = client.ListMessages(client.MailboxInfo.InboxUri);

                // Loop through the collection to display the basic information
                foreach (ExchangeMessageInfo msgInfo in msgCollection)
                {
                    string strMsgInfo = "Subject: " + msgInfo.Subject + " == From: " + msgInfo.From.ToString() + " == To: " + msgInfo.To.ToString();
                    lstMessages.Items.Add(strMsgInfo);
                }
                // ExEnd:AsposeEmailExchange

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
