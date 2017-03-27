using System;
using System.Collections.Generic;
using System.Configuration;
using System.Net;
using System.Threading;
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
    class CreateFoldersOnExchangeServerMailbox
    {
        public static void Run()
        {
            // ExStart:CreateFoldersOnExchangeServerMailbox

            // Create instance of EWSClient class by giving credentials
            IEWSClient client = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain");

            string inbox = client.MailboxInfo.InboxUri;
            string folderName1 = "EMAILNET-35054";
            string subFolderName0 = "2015";
            string folderName2 = folderName1 + "/" + subFolderName0;
            string folderName3 = folderName1 + " / 2015";
            ExchangeFolderInfo rootFolderInfo = null;
            ExchangeFolderInfo folderInfo = null;

            try
            {
                client.UseSlashAsFolderSeparator = true;
                client.CreateFolder(client.MailboxInfo.InboxUri, folderName1);
                client.CreateFolder(client.MailboxInfo.InboxUri, folderName2);
            }
            finally
            {
                if (client.FolderExists(inbox, folderName1, out rootFolderInfo))
                {
                    if (client.FolderExists(inbox, folderName2, out folderInfo))
                        client.DeleteFolder(folderInfo.Uri, true);
                    client.DeleteFolder(rootFolderInfo.Uri, true);
                }
            }
            // ExEnd:CreateFoldersOnExchangeServerMailbox
        }
    }
}