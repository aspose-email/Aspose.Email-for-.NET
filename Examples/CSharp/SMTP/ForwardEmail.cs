using System;
using System.IO;
using System.Security.Cryptography;
using Aspose.Email.Mime;
using Aspose.Email.Clients.Smtp;
using Aspose.Email.Clients;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.SMTP
{
    class ForwardEmail
    {
        public static void Run()
        {
            // ExStart:ForwardEmail
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_SMTP();

            //Create an instance of SmtpClient class
            SmtpClient client = new SmtpClient();

            // Specify your mailing host server, Username, Password, Port and SecurityOptions
            client.Host = "mail.server.com";
            client.Username = "username";
            client.Password = "password";
            client.Port = 587;
            client.SecurityOptions = SecurityOptions.SSLExplicit;
            MailMessage message = MailMessage.Load(dataDir + "Message.eml");
            client.Forward("Recipient1@domain.com", "Recipient2@domain.com", message);
            // ExEnd:ForwardEmail
        }
    }
}