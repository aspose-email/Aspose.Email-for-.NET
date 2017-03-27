using System;
using System.Net;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Mime;
using Aspose.Email.Mapi;
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
    class DownloadMessagesFromPublicFolders
    {
        // ExStart:DownloadMessagesFromPublicFolders
        public static string dataDir = RunExamples.GetDataDir_Exchange();
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
                msg.Save(dataDir +  msg.Subject + ".msg",  SaveOptions.DefaultMsgUnicode);
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
        // ExEnd:DownloadMessagesFromPublicFolders
    }
}