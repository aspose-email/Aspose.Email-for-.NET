using System;
using Aspose.Email.Clients;
using Aspose.Email.Clients.Base;
using Aspose.Email.Clients.Imap;


namespace Aspose.Email.Examples.CSharp.Email.IMAP
{
    class ChangeOrderOfEmails
    {
        public static void Run()
        {
            // ExStart:ChangeOrderOfEmails
            ImapClient imapClient = new ImapClient();
            imapClient.Host = "<HOST>";
            imapClient.Port = 993;
            imapClient.Username = "<USERNAME>";
            imapClient.Password = "<PASSWORD>";
            imapClient.SupportedEncryption = EncryptionProtocols.Tls;
            imapClient.SecurityOptions = SecurityOptions.SSLImplicit;

            PageSettings pageSettings = new PageSettings { AscendingSorting = true };
            ImapPageInfo pageInfo = imapClient.ListMessagesByPage(5, pageSettings);
            ImapMessageInfoCollection messages = pageInfo.Items;

            foreach (ImapMessageInfo message in messages)
            {
                Console.WriteLine(message.Subject + " -> " + message.Date.ToString());
            }
            // ExEnd:ChangeOrderOfEmails
        }
    }
}