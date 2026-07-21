using System;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Clients;

namespace Aspose.Email.Examples.EWS
{
    class ConnectExchangeServerUsingIMAP
    {
        public static void Run()
        {
            using (var imapClient = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                imapClient.SecurityOptions = SecurityOptions.Auto;

                // Select the Inbox folder
                imapClient.SelectFolder(ImapFolderInfo.InBox);

                // Get the list of messages
                var msgCollection = imapClient.ListMessages();

                foreach (var msgInfo in msgCollection)
                {
                    Console.WriteLine(msgInfo.Subject);
                }
            }
        }
    }
}