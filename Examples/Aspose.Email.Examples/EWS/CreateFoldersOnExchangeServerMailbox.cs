using System;
using System.Collections.Generic;
using System.Configuration;
using System.Net;
using System.Threading;
using Aspose.Email.Mime;
using Aspose.Email.Mapi;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Clients.Exchange;

namespace Aspose.Email.Examples.EWS
{
    class CreateFoldersOnExchangeServerMailbox
    {
        public static void Run()
        {

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                client.UseSlashAsFolderSeparator = true;

                client.CreateFolder(client.MailboxInfo.InboxUri, "Test Folder/Test Subfolder");
                client.CreateFolder(client.MailboxInfo.InboxUri, "Level1/Level2/Level3");
            }
        }
    }
}