using Aspose.Email.Clients.Exchange;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Common.Delegate;
using Aspose.Email.Mapi;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;

namespace CSharp.Exchange_EWS
{
    public class MoveItemsToInPlaceArchive
    {
        public static void Run()
        {
            //ExStart: MoveItemsToInPlaceArchive
            const string mailboxUri = "<HOST>";
            const string domain = @"";
            const string username = @"<USERNAME>";
            const string password = @"<PASSWORD>";
            NetworkCredential credentials = new NetworkCredential(username, password, domain);
            IEWSClient client = EWSClient.GetEWSClient(mailboxUri, credentials);

            ExchangeMessageInfoCollection msgCollection = client.ListMessages(client.MailboxInfo.InboxUri);

            foreach (ExchangeMessageInfo msgInfo in msgCollection)
            {
                Console.WriteLine("Subject:" + msgInfo.Subject);
                client.ArchiveItem(client.MailboxInfo.InboxUri, msgInfo.UniqueUri);
            }
            client.Dispose();
            //ExEnd: MoveItemsToInPlaceArchive
        }
    }
}
