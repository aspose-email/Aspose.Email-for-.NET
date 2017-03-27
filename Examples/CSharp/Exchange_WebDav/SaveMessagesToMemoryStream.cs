using System;
using System.IO;
using Aspose.Email.Clients.Exchange.Dav;
using Aspose.Email.Clients.Exchange;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Exchange_WebDav
{
    class SaveMessagesToMemoryStream
    {
        public static void Run()
        {
            try
            {
                // ExStart:SaveMessagesToMemoryStream
                string dataDir = RunExamples.GetDataDir_Exchange();
                // Create instance of ExchangeClient class by giving credentials
                ExchangeClient client = new ExchangeClient("https://Ex07sp1/exchange/Administrator", "user", "pwd",
                    "domain");

                // Call ListMessages method to list messages info from Inbox
                ExchangeMessageInfoCollection msgCollection = client.ListMessages(client.MailboxInfo.InboxUri);

                // Loop through the collection to get Message URI
                foreach (ExchangeMessageInfo msgInfo in msgCollection)
                {
                    string strMessageURI = msgInfo.UniqueUri;

                    // Now save the message in memory stream
                    MemoryStream stream = new MemoryStream();
                    client.SaveMessage(strMessageURI, stream);
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