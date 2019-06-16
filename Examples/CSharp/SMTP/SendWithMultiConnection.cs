using System;
using System.Collections.Generic;
using Aspose.Email.Clients;
using Aspose.Email.Clients.Base;
using Aspose.Email.Clients.Smtp;

namespace Aspose.Email.Examples.CSharp.Email.IMAP
{
    class SendWithMultiConnection
    {
        public static void Run()
        {
            //ExStart: 1
            SmtpClient smtpClient = new SmtpClient();
            smtpClient.Host = "<HOST>";
            smtpClient.Username = "<USERNAME>";
            smtpClient.Password = "<PASSWORD>";
            smtpClient.Port = 587;
            smtpClient.SupportedEncryption = EncryptionProtocols.Tls;
            smtpClient.SecurityOptions = SecurityOptions.SSLExplicit;

            List<MailMessage> messages = new List<MailMessage>();
            for (int i = 0; i < 20; i++)
            {
                MailMessage message = new MailMessage(
                    "<EMAIL ADDRESS>",
                    "<EMAIL ADDRESS>",
                    "Test Message - " + Guid.NewGuid().ToString(),
                    "SMTP Send Messages with MultiConnection");
                messages.Add(message);
            }

            smtpClient.ConnectionsQuantity = 5;
            smtpClient.UseMultiConnection = MultiConnectionMode.Enable;
            smtpClient.Send(messages);
            //ExEnd: 1

            Console.WriteLine("SendWithMultiConnection executed successfully.");
        }
    }
}
