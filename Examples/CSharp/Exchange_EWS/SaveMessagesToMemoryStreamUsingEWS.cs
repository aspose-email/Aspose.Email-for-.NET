using System;
using System.IO;
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
    class SaveMessagesToMemoryStreamUsingEWS
    {
        public static void Run()
        {

            try
            {
                // ExStart:SaveMessagesToMemoryStream
                string datadir = RunExamples.GetDataDir_Exchange();
                // Create instance of EWSClient class by giving credentials
                IEWSClient client = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain");

                // Call ListMessages method to list messages info from Inbox
                ExchangeMessageInfoCollection msgCollection = client.ListMessages(client.MailboxInfo.InboxUri);

                // Loop through the collection to get Message URI
                foreach (ExchangeMessageInfo msgInfo in msgCollection)
                {
                    string strMessageURI = msgInfo.UniqueUri;

                    // Now save the message in memory stream
                    MemoryStream stream = new MemoryStream();
                    client.SaveMessage(strMessageURI, datadir + stream);
                }
                // ExEnd:SaveMessagesToMemoryStream
            }
            catch (Exception ex)
            {

                Console.Write(ex.Message);
            }
        }
    }
}