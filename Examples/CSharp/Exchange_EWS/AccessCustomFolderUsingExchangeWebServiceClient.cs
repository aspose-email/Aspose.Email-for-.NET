using System;
using Aspose.Email.Mime;
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
    class AccessCustomFolderUsingExchangeWebServiceClient
    {
        public static void Run()
        {
            // ExStart:AccessCustomFolderUsingExchangeWebServiceClient
            // Create instance of EWSClient class by giving credentials
            IEWSClient client = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain");

            // Create ExchangeMailboxInfo, ExchangeMessageInfoCollection instance
            ExchangeMailboxInfo mailbox = client.GetMailboxInfo();
            ExchangeMessageInfoCollection messages = null;
            ExchangeFolderInfo subfolderInfo = new ExchangeFolderInfo();

            // Check if specified custom folder exisits and Get all the messages info from the target Uri
            client.FolderExists(mailbox.InboxUri, "TestInbox", out subfolderInfo);

            //if custom folder exists
            if (subfolderInfo != null)
            {
                messages = client.ListMessages(subfolderInfo.Uri);

                // Parse all the messages info collection
                foreach (ExchangeMessageInfo info in messages)
                {
                    string strMessageURI = info.UniqueUri;
                    // now get the message details using FetchMessage()
                    MailMessage msg = client.FetchMessage(strMessageURI);
                    Console.WriteLine("Subject: " + msg.Subject);
                }
            }
            else
            {
                Console.WriteLine("No folder with this name found.");
            }
            // ExEnd:AccessCustomFolderUsingExchangeWebServiceClient
        }
    }
}