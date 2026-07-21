using Aspose.Email.Clients.Exchange;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Mime;

namespace Aspose.Email.Examples.EWS
{
    class AccessAnotherMailboxUsingExchangeWebServiceClient
    {
        public static void Run()
        {
            // Create instance of EWSClient class by giving credentials
            IEWSClient client = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain");

            // Get Exchange mailbox info of other email account
            ExchangeMailboxInfo mailboxInfo = client.GetMailboxInfo("otherUser@domain.com");
        }
    }
}