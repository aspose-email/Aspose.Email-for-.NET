using System;
using System.Collections.Generic;
using System.Net;
using Aspose.Email.Mime;
using Aspose.Email.Mapi;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Clients.Exchange;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Exchange_EWS
{
    class SynchronizeFolderItems
    {
        public static void Run()
        {
            try
            {

                // ExStart:SynchronizeFolderItems
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
                // ExEnd:SynchronizeFolderItems
            }
            catch (Exception ex)
            {

                Console.Write(ex.Message);
            }
        }
    }
}