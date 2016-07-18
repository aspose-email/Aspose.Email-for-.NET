using System;
using System.IO;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using Aspose.Email.Exchange;
using Aspose.Email.Mail;
using Aspose.Email.Outlook;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Exchange
{
    class DownloadMessagesFromExchangeServerFoldersRecursively
    {
        static string mailboxURI = "https:// Ex2010/ews/exchange.asmx"; // EWS

        static string username = "administrator";
        static string password = "pwd";
        static string domain = "ex2010.local";

        public static void Run()
        {
            // ExStart:DeleteMessagesFromExchangeServer
            // Create instance of IEWSClient class by giving credentials
            string mailboxURI = "http:// Ex2003/exchange/administrator"; // WebDAV 

            // Register callback method for SSL validation event
            ServicePointManager.ServerCertificateValidationCallback += RemoteCertificateValidationHandler;

            try
            {

                DownloadAllMessages();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            // ExEnd:DeleteMessagesFromExchangeServer
        }

        private static void DownloadAllMessages()
        {
            try
            {
                string rootFolder = domain + "-" + username;
                Directory.CreateDirectory(rootFolder);
                string inboxFolder = rootFolder + "\\Inbox";
                Directory.CreateDirectory(inboxFolder);

                Console.WriteLine("Downloading all messages from Inbox....");
                // Create instance of IEWSClient class by giving credentials
                IEWSClient client = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", username, password, domain);

                ExchangeMailboxInfo mailboxInfo = client.GetMailboxInfo();
                Console.WriteLine("Mailbox URI: " + mailboxInfo.MailboxUri);
                string rootUri = client.GetMailboxInfo().RootUri;
                // List all the folders from Exchange server
                ExchangeFolderInfoCollection folderInfoCollection = client.ListSubFolders(rootUri);
                foreach (ExchangeFolderInfo folderInfo in folderInfoCollection)
                {
                    // Call the recursive method to read messages and get sub-folders
                    ListMessagesInFolder(client, folderInfo, rootFolder);
                }

                Console.WriteLine("All messages downloaded.");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        /// <summary>
        /// Recursive method to get messages from folders and sub-folders
        /// </summary>
        /// <param name="folderInfo"></param>
        /// <param name="rootFolder"></param>
        private static void ListMessagesInFolder(IEWSClient client, ExchangeFolderInfo folderInfo, string rootFolder)
        {
            // Create the folder in disk (same name as on IMAP server)
            string currentFolder = rootFolder + "\\" + folderInfo.DisplayName;
            Directory.CreateDirectory(currentFolder);

            // List messages
            ExchangeMessageInfoCollection msgInfoColl = client.ListMessages(folderInfo.Uri);
            Console.WriteLine("Listing messages....");
            int i = 0;
            foreach (ExchangeMessageInfo msgInfo in msgInfoColl)
            {
                // Get subject and other properties of the message
                Console.WriteLine("Subject: " + msgInfo.Subject);

                // Get rid of characters like ? and :, which should not be included in a file name
                string fileName = msgInfo.Subject.Replace(":", " ").Replace("?", " ");

                // Save the message in MSG format
                MailMessage msg = client.FetchMessage(msgInfo.UniqueUri);
                // Msg.Save(msg);
                i++;
            }
            Console.WriteLine("============================\n");
            try
            {
                // If this folder has sub-folders, call this method recursively to get messages
                ExchangeFolderInfoCollection folderInfoCollection = client.ListSubFolders(folderInfo.Uri);
                foreach (ExchangeFolderInfo subfolderInfo in folderInfoCollection)
                {
                    ListMessagesInFolder(client, subfolderInfo, currentFolder);
                }
            }
            catch (Exception)
            {
            }
        }

        private static bool RemoteCertificateValidationHandler(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors)
        {
            return true; // Ignore the checks and go ahead
        }
    }
}