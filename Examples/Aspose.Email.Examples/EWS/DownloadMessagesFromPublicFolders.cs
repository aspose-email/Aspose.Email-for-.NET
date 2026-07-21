using System;
using System.Net;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Mime;
using Aspose.Email.Mapi;
using Aspose.Email.Clients.Exchange;

namespace Aspose.Email.Examples.EWS
{
    class DownloadMessagesFromPublicFolders
    {
        public static string mailboxUri = "https://exchange/ews/exchange.asmx"; // EWS
        public static string username = "administrator";
        public static string password = "pwd";
        public static string domain = "ex2013.local";

        public static void Run()
        {
            try
            {
                ReadPublicFolders();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private static void ReadPublicFolders()
        {
            NetworkCredential credential = new NetworkCredential(username, password, domain);
            IEWSClient client = EWSClient.GetEWSClient(mailboxUri, credential);

            ExchangeFolderInfoCollection folders = client.ListPublicFolders();
            foreach (ExchangeFolderInfo publicFolder in folders)
            {
                Console.WriteLine("Name: " + publicFolder.DisplayName);
                Console.WriteLine("Subfolders count: " + publicFolder.ChildFolderCount);
                ListMessagesFromSubFolder(publicFolder, client);

            }
        }

        private static void ListMessagesFromSubFolder(ExchangeFolderInfo publicFolder, IEWSClient client)
        {
            Console.WriteLine("Folder Name: " + publicFolder.DisplayName);
            ExchangeMessageInfoCollection msgInfoCollection = client.ListMessagesFromPublicFolder(publicFolder);
            foreach (ExchangeMessageInfo messageInfo in msgInfoCollection)
            {
                MailMessage msg = client.FetchMessage(messageInfo.UniqueUri);
                Console.WriteLine(msg.Subject);
                msg.Save(Data.Out +  msg.Subject + ".msg",  SaveOptions.DefaultMsgUnicode);
            }

            // Call this method recursively for any subfolders
            if (publicFolder.ChildFolderCount > 0)
            {
                ExchangeFolderInfoCollection subfolders = client.ListSubFolders(publicFolder);
                foreach (ExchangeFolderInfo subfolder in subfolders)
                {
                    ListMessagesFromSubFolder(subfolder, client);
                }
            }
        }
    }
}