using System;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Mime;
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
    class AccessAnotherMailboxUsingExchangeClient
    {
        public static void Run()
        {
            // ExStart:AccessAnotherMailboxUsingExchangeClient
            // Create instance of ExchangeClient class by giving credentials
            ExchangeClient client = new ExchangeClient("http://MachineName/exchange/Username","Username", "password", "domain");

            // Get Exchange mailbox info of other email account
            ExchangeMailboxInfo mailboxInfo = client.GetMailboxInfo("otherUser@domain.com");
            // ExEnd:AccessAnotherMailboxUsingExchangeClient
        }
    }
}