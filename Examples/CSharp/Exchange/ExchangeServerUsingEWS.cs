using System;
using Aspose.Email.Exchange;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Exchange
{
    class ExchangeServerUsingEWS
    {
        public static void Run()
        {
            try
            {
                // ExStart:ExchangeServerUsingEWS
                // Create instance of EWSClient class by giving credentials
                IEWSClient client = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain");

                // Get folder URI
                string strFolderURI = string.Empty;
                strFolderURI = client.MailboxInfo.InboxUri;
                strFolderURI = client.MailboxInfo.DeletedItemsUri;
                strFolderURI = client.MailboxInfo.DraftsUri;
                strFolderURI = client.MailboxInfo.SentItemsUri;

                // Get list of messages from the specified folder
                ExchangeMessageInfoCollection msgCollection = client.ListMessages(strFolderURI);
                // ExEnd:ExchangeServerUsingEWS
            }
            catch (Exception ex)
            {

                Console.Write(ex.Message);
            }
        }
    }
}