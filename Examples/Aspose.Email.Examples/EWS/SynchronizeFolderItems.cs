using System;
using System.Collections.Generic;
using System.Net;
using Aspose.Email.Mime;
using Aspose.Email.Mapi;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Clients.Exchange;

namespace Aspose.Email.Examples.EWS
{
    class SynchronizeFolderItems
    {
        public static void Run()
        {
            try
            {

                IEWSClient client = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain");
                MailMessage message1 = new MailMessage("user@domain.com", "user@domain.com", "EMAILNET-34738 - " + Guid.NewGuid().ToString(),  "EMAILNET-34738 Sync Folder Items");
                client.Send(message1);

                MailMessage message2 = new MailMessage("user@domain.com", "user@domain.com", "EMAILNET-34738 - " + Guid.NewGuid().ToString(),"EMAILNET-34738 Sync Folder Items");
                client.Send(message2);

                ExchangeMessageInfoCollection messageInfoCol = client.ListMessages(client.MailboxInfo.InboxUri);
                SyncFolderResult result = client.SyncFolder(client.MailboxInfo.InboxUri, null);
                Console.WriteLine(result.NewItems.Count);
                Console.WriteLine(result.ChangedItems.Count);
                Console.WriteLine(result.ReadFlagChanged.Count);
                Console.WriteLine(result.DeletedItems.Length);
            }
            catch (Exception ex)
            {

                Console.Write(ex.Message);
            }
        }
    }
}