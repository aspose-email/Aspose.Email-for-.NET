using Aspose.Email.Clients;
using Aspose.Email.Clients.Smtp;
using Aspose.Email.Mime;
using System;
using System.Collections.Generic;
using System.Diagnostics;
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
    class SendEmailsSynchronously
    {
        public static void Run()
        {
            // ExStart:SendEmailsSynchronously
            // Declare msg as MailMessage instance
            MailMessage msg = new MailMessage();

            // Create an instance of SmtpClient class
            SmtpClient client = new SmtpClient();

            // Specify your mailing host server, Username, Password, Port # and Security option
            client.Host = "mail.server.com";
            client.Username = "username";
            client.Password = "password";
            client.Port = 587;
            client.SecurityOptions = SecurityOptions.SSLExplicit;

            try
            {
                // Client.Send will send this message
                client.Send(msg);
                Console.WriteLine("Message sent");
            }
            catch (Exception ex)
            {
                Trace.WriteLine(ex.ToString());
            }
            // ExEnd:SendEmailsSynchronously
        }
    }
}
