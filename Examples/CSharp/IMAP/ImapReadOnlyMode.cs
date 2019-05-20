using Aspose.Email;
using Aspose.Email.Clients;
using Aspose.Email.Clients.Base;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Tools.Search;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CSharp.IMAP
{
    public class ImapReadOnlyMode
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

            ImapQueryBuilder imapQueryBuilder = new ImapQueryBuilder();
            imapQueryBuilder.HasNoFlags(ImapMessageFlags.IsRead); /* get unread messages. */
            MailQuery query = imapQueryBuilder.GetQuery();

            imapClient.ReadOnly = true;
            imapClient.SelectFolder("Inbox");
            ImapMessageInfoCollection messageInfoCol = imapClient.ListMessages(query);
            Console.WriteLine("Initial Unread Count: " + messageInfoCol.Count());
            if (messageInfoCol.Count() > 0)
            {
                imapClient.FetchMessage(messageInfoCol[0].SequenceNumber);

                messageInfoCol = imapClient.ListMessages(query);
                // This count will be equal to the initial count
                Console.WriteLine("Updated Unread Count: " + messageInfoCol.Count());
            }
            else
            {
                Console.WriteLine("No unread messages found");
            }
            //ExEnd: 1

            Console.WriteLine("ImapReadOnlyMode executed successfully.");
        }
    }
}
