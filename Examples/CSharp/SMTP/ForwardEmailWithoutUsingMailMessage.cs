using Aspose.Email;
using Aspose.Email.Clients;
using Aspose.Email.Clients.Smtp;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.SMTP
{
    class ForwardEmailWithoutUsingMailMessage
    {
        public static void Run()
        {
            //ExStart: ForwardEmailWithoutUsingMailMessage
            string dataDir = RunExamples.GetDataDir_Email();

            string host = "mail.server.com";
            string username = "username";
            string password = "password";
            int smtpPort = 587;
            string sender = "Sender@domain.com";
            MailAddressCollection recipients = new MailAddressCollection();
            recipients.Add("recepient1@domain.com, recepient2@domain.com");

            using (SmtpClient client = new SmtpClient(host, smtpPort, username, password, SecurityOptions.Auto))
            {
                string fileName = @"test.eml";
                using (FileStream fs = File.OpenRead(dataDir + fileName))
                {
                    client.Forward(sender, recipients, fs);
                }
            }
            //ExEnd: ForwardEmailWithoutUsingMailMessage
        }
    }
}
