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
    class ListFoldersFromExchangeServer
    {
        public static void Run()
        {
            try
            {
                IEWSClient client = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain");
                Console.WriteLine("Downloading all messages from Inbox....");

                ExchangeMailboxInfo mailboxInfo = client.GetMailboxInfo();
                Console.WriteLine("Mailbox URI: " + mailboxInfo.MailboxUri);
                string rootUri = client.GetMailboxInfo().RootUri;
                // List all the folders from Exchange server
                ExchangeFolderInfoCollection folderInfoCollection = client.ListSubFolders(rootUri);
                foreach (ExchangeFolderInfo folderInfo in folderInfoCollection)
                {
                    // Call the recursive method to read messages and get sub-folders
                    ListSubFolders(client, folderInfo);
                }

                Console.WriteLine("All messages downloaded.");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        private static void ListSubFolders(IEWSClient client, ExchangeFolderInfo folderInfo)
        {
            // Create the folder in disk (same name as on IMAP server)
            Console.WriteLine(folderInfo.DisplayName);
            try
            {
                // If this folder has sub-folders, call this method recursively to get messages
                ExchangeFolderInfoCollection folderInfoCollection = client.ListSubFolders(folderInfo.Uri);
                foreach (ExchangeFolderInfo subfolderInfo in folderInfoCollection)
                {
                    ListSubFolders(client, subfolderInfo);
                }
            }
            catch (Exception)
            {
            }
        }
    }
}