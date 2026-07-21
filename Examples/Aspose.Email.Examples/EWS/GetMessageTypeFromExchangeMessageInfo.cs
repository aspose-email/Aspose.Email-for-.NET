using Aspose.Email.Clients.Exchange;
using Aspose.Email.Clients.Exchange.WebService;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;

namespace Aspose.Email.Examples.EWS
{
    class GetMessageTypeFromExchangeMessageInfo
    {
        static void Run()
        {
            const string mailboxUri = "https://exchange/ews/exchange.asmx";
            const string domain = @"";
            const string username = @"username@ASE305.onmicrosoft.com";
            const string password = @"password";
            NetworkCredential credentials = new NetworkCredential(username, password, domain);

            IEWSClient client = EWSClient.GetEWSClient(mailboxUri, credentials);

            ExchangeMessageInfoCollection list = client.ListMessages(client.MailboxInfo.DeletedItemsUri);
            Console.WriteLine(list[0].MessageInfoType.ToString());
        }
    }
}
