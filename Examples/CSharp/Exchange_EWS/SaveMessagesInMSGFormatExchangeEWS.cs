using Aspose.Email.Clients.Exchange;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Mime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Exchange_EWS
{
    class SaveMessagesInMSGFormatExchangeEWS
    {
        public static void Run()
        {
            string dataDir = RunExamples.GetDataDir_Exchange();

            try
            {
                // ExStart:SaveMessagesInMSGFormatExchangeEWS
                // Create instance of EWSClient class by giving credentials
                IEWSClient client = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain");

                // Call ListMessages method to list messages info from Inbox
                ExchangeMessageInfoCollection msgCollection = client.ListMessages(client.MailboxInfo.InboxUri);

                int count = 0;
                // Loop through the collection to get Message URI
                foreach (ExchangeMessageInfo msgInfo in msgCollection)
                {
                    string strMessageURI = msgInfo.UniqueUri;
                    
                    // Now get the message details using FetchMessage() and Save message as Msg 
                    MailMessage message = client.FetchMessage(strMessageURI);
                    message.Save(dataDir + (count++) + "_out.msg", SaveOptions.DefaultMsgUnicode);
                }
                // ExEnd:SaveMessagesInMSGFormatExchangeEWS
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
