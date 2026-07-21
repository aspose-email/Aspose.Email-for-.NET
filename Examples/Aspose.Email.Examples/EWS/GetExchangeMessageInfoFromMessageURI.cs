using Aspose.Email.Clients.Exchange;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Mime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Aspose.Email.Examples.EWS
{
    class GetExchangeMessageInfoFromMessageURI
    {
        public static void Run()
        {
            try
            {
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
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }            
        }
    }
}
