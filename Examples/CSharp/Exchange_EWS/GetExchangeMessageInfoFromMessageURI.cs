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
    class GetExchangeMessageInfoFromMessageURI
    {
        public static void Run()
        {
            try
            {
                // ExStart:GetExchangeMessageInfoFromMessageURI
                IEWSClient client = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "user@domain.com", "pwd", "domain");

                List<string> ids = new List<string>();
                List<MailMessage> messages = new List<MailMessage>();

                for (int i = 0; i < 5; i++)
                {
                    MailMessage message = new MailMessage("user@domain.com", "receiver@domain.com", "EMAILNET-35033 - " + Guid.NewGuid().ToString(), "EMAILNET-35033 Messages saved from Sent Items folder doesn't contain 'To' field");
                    messages.Add(message);
                    string uri = client.AppendMessage(message);
                    ids.Add(uri);
                }

                ExchangeMessageInfoCollection messageInfoCol = client.ListMessages(ids);

                foreach (ExchangeMessageInfo messageInfo in messageInfoCol)
                {
                    // Do something ...
                    Console.WriteLine(messageInfo.UniqueUri);
                }
                // ExEnd:GetExchangeMessageInfoFromMessageURI
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }            
        }
    }
}
