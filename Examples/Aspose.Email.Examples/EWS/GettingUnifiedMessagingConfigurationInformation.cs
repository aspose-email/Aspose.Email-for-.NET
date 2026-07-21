using Aspose.Email.Clients.Exchange.WebService;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;

namespace Aspose.Email.Examples.EWS
{
    class GettingUnifiedMessagingConfigurationInformation
    {
        public static void Run()
        {
            const string mailboxUri = "https://exchange.domain.com/ews/Exchange.asmx";
            const string domain = @"";
            const string username = @"username";
            const string password = @"password";
            NetworkCredential credential = new NetworkCredential(username, password, domain);

            try
            {
                IEWSClient client = EWSClient.GetEWSClient(mailboxUri, credential);
                UnifiedMessagingConfiguration umConf = client.GetUMConfiguration();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
