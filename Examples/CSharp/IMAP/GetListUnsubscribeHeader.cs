using Aspose.Email.Clients;
using Aspose.Email.Clients.Base;
using Aspose.Email.Clients.Imap;
using System;

namespace CSharp.IMAP
{
    public class GetListUnsubscribeHeader
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

            ImapMessageInfoCollection messageInfoCol = imapClient.ListMessages();
            foreach (ImapMessageInfo imapMessageInfo in messageInfoCol)
            {
                Console.WriteLine("ListUnsubscribe Header: " + imapMessageInfo.ListUnsubscribe);
            }
            //ExEnd: 1

            Console.WriteLine("GetListUnsubscribeHeader executed successfully.");
        }
    }
}
