using Aspose.Email.Clients;
using Aspose.Email.Clients.Pop3;
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
    public partial class AsposeEmailPop3SampleApp : Form
    {
        public AsposeEmailPop3SampleApp()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // ExStart:AsposeEmailPop3
            // Create a POP3 client
            Pop3Client client = new Pop3Client();

            // Basic settings (required)
            client.Host = "pop3.youdomain.com";
            client.Username = "username";
            client.Password = "psw";

            try
            {
                // Retrieve first message in MailMessage format directly
                Aspose.Email.MailMessage msg;
                msg = client.FetchMessage(1);
                txtFrom.Text = msg.From.ToString();
                txtSubject.Text = msg.Subject.ToString();
                txtBody.Text = msg.HtmlBody.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            // ExEnd:AsposeEmailPop3
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // Create a POP3 client
            Pop3Client client = new Pop3Client();

            // Basic settings (required)
            client.Host = "pop3.youdomain.com";
            client.Username = "username";
            client.Password = "psw";

            // ExStart:SSLEnabledServer
            // Set implicit security mode
            client.SecurityOptions = SecurityOptions.SSLImplicit;
            // ExEnd:SSLEnabledServer

            try
            {
                // Retrieve first message in MailMessage format directly
                Aspose.Email.MailMessage msg;
                msg = client.FetchMessage(1);
                txtFrom.Text = msg.From.ToString();
                txtSubject.Text = msg.Subject.ToString();
                txtBody.Text = msg.HtmlBody.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
