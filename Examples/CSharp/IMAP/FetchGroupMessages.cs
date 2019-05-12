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
    public class FetchGroupMessages
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
            Console.WriteLine("ListMessages Count: " + messageInfoCol.Count);
            int[] sequenceNumberAr = messageInfoCol.Select((ImapMessageInfo mi) => mi.SequenceNumber).ToArray();
            string[] uniqueIdAr = messageInfoCol.Select((ImapMessageInfo mi) => mi.UniqueId).ToArray();

            IList<MailMessage> fetchedMessagesBySNumMC = imapClient.FetchMessages(sequenceNumberAr);
            Console.WriteLine("FetchMessages-sequenceNumberAr Count: " + messageInfoCol.Count);

            IList<MailMessage> fetchedMessagesByUidMC = imapClient.FetchMessages(uniqueIdAr);
            Console.WriteLine("FetchMessages-uniqueIdAr Count: " + messageInfoCol.Count);
            //ExEnd: 1

            Console.WriteLine("FetchGroupMessages executed successfully.");
        }
    }
}
