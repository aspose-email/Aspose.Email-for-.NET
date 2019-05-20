using Aspose.Email;
using Aspose.Email.Clients;
using Aspose.Email.Clients.Base;
using Aspose.Email.Clients.Imap;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CSharp.IMAP
{
    public class ImapGroupAppendWithMultiConnection
    {
        public static void Run()
        {
            //ExStart: 1
            ImapClient imapClient = new ImapClient();
            imapClient.Host = "<HOST>";
            imapClient.Port = 993;
            imapClient.Username = "<USERNAME>";
            imapClient.Password = "<PASSWORD>";
            imapClient.SupportedEncryption = EncryptionProtocols.Tls;
            imapClient.SecurityOptions = SecurityOptions.SSLImplicit;

            List<MailMessage> messages = new List<MailMessage>();
            for (int i = 0; i < 20; i++)
            {
                MailMessage message = new MailMessage(
                    "<EMAIL ADDRESS>",
                    "<EMAIL ADDRESS>",
                    "Test Message - " + Guid.NewGuid().ToString(),
                    "IMAP Group Append with MultiConnection");
                messages.Add(message);
            }

            imapClient.ConnectionsQuantity = 5;
            imapClient.UseMultyConnection = MultyConnectionMode.Enable;
            imapClient.AppendMessages(messages);
            //ExEnd: 1

            Console.WriteLine("ImapGroupAppendWithMultiConnection executed successfully.");
        }
    }
}
