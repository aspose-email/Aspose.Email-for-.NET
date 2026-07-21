using Aspose.Email.Clients.Exchange;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Mime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;

namespace Aspose.Email.Examples.EWS
{
    class SendEmailToPrivateDistributionList
    {
        public static void Run()
        {
            try
            {
                // Set mailboxURI, Username, password, domain information
                string mailboxUri = "https://ex2010/ews/exchange.asmx";
                string username = "test.exchange";
                string password = "pwd";
                string domain = "ex2010.local";
                NetworkCredential credentials = new NetworkCredential(username, password, domain);
                IEWSClient client = EWSClient.GetEWSClient(mailboxUri, credentials);

                ExchangeDistributionList[] distributionLists = client.ListDistributionLists();
                MailAddress distributionListAddress = distributionLists[0].ToMailAddress();
                MailMessage message = new MailMessage(new MailAddress("from@host.com"), distributionListAddress);
                message.Subject = "sendToPrivateDistributionList";
                client.Send(message);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
