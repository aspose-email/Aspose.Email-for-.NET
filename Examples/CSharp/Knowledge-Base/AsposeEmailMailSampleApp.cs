using Aspose.Email.Clients;
using Aspose.Email.Clients.Smtp;
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
    public partial class AsposeEmailMailSampleApp : Form
    {
        public AsposeEmailMailSampleApp()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                // ExStart:AsposeEmailMail
                // Declare message as MailMessage instance
                MailMessage message = new MailMessage();
                // Specify From, To, Subject and Body 
                message.From = txtFrom.Text;
                message.To = txtTo.Text;
                message.Subject = txtSubject.Text;
                message.Body = txtBody.Text;

                // Send email using SmtpClient, Create an instance of the SmtpClient Class and Specify the mailing host server, Username, Password and Port
                SmtpClient client = new SmtpClient();

                // Specify the mailing host server, Username, Password and Port 
                client.Host = "mail.server.com";
                client.Username = "username";
                client.Password = "password";
                client.Port = 25;
                client.Send(message);

                // Notify the user that a message has been sent
                MessageBox.Show("Message Sent");
                // ExEnd:AsposeEmailMail

                // ExStart:SSLEnabledSMTP
                // Set the port to 587. This is the SSL port of Gmail SMTP server, set the security mode to explicit
                client.Port = 587;
                client.SecurityOptions = SecurityOptions.SSLExplicit;
                // ExEnd:SSLEnabledSMTP

                client.Send(message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
