using Aspose.Email.Clients.Exchange;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Mime;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Exchange_EWS
{
    class AccessAnotherMailboxUsingExchangeWebServiceClient
    {
        public static void Run()
        {
            // ExStart:AccessAnotherMailboxUsingExchangeWebServiceClient
            // Create instance of EWSClient class by giving credentials
            IEWSClient client = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain");

            // Get Exchange mailbox info of other email account
            ExchangeMailboxInfo mailboxInfo = client.GetMailboxInfo("otherUser@domain.com");
            // ExEnd:AccessAnotherMailboxUsingExchangeWebServiceClient
        }
    }
}