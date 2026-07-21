using System;
using System.IO;
using System.Security.Cryptography;
using Aspose.Email.Mime;
using Aspose.Email.Clients.Smtp;
using Aspose.Email.Clients;

namespace Aspose.Email.Examples.SMTP
{
    class ForwardEmail
    {
        public static void Run()
        {

            //Create an instance of SmtpClient class
            SmtpClient client = new SmtpClient();

            // Specify your mailing host server, Username, Password, Port and SecurityOptions
            client.Host = "mail.server.com";
            client.Username = "username";
            client.Password = "password";
            client.Port = 587;
            client.SecurityOptions = SecurityOptions.SSLExplicit;
            MailMessage message = MailMessage.Load(Data.Smtp + "Message.eml");
            client.Forward("Recipient1@domain.com", "Recipient2@domain.com", message);
        }
    }
}