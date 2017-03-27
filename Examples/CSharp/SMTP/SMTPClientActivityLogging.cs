using Aspose.Email.Clients;
using Aspose.Email.Clients.Smtp;
using Aspose.Email.Mime;
using System;
using System.Collections.Generic;
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
    class SMTPClientActivityLogging
    {
        public static void Run()
        {
            // ExStart:SMTPClientActivityLogging
            // Build message
            MailMessage message = new MailMessage();

            // Set email address for From and TO
            message.From = "userFrom@gmail.com";
            message.To = "userTo@gmail.com";

            // Set Subject and Body
            message.Subject = "Appointment Request";
            message.Body = "Test Body";

            // Initialize SmtpClient and Set valid user name and password, Port and SecurityOptions
            SmtpClient client = new SmtpClient();
            client.Host = "smtp.gmail.com";
            client.Username = "userFrom";
            client.Password = "***********";
            client.Port = 587;
            client.SecurityOptions = SecurityOptions.SSLExplicit;
            try
            {
                client.Send(message);
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            // ExEnd:SMTPClientActivityLogging
        }
    }
}
