using Aspose.Email.Clients.Exchange;
using Aspose.Email.Clients.Exchange.WebService;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
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
    class GetMessageTypeFromExchangeMessageInfo
    {
        static void Run()
        {
            //ExStart: GetMessageTypeFromExchangeMessageInfo
            const string mailboxUri = "https://exchange/ews/exchange.asmx";
            const string domain = @"";
            const string username = @"username@ASE305.onmicrosoft.com";
            const string password = @"password";
            NetworkCredential credentials = new NetworkCredential(username, password, domain);

            // ExStart:CopyConversations
            IEWSClient client = EWSClient.GetEWSClient(mailboxUri, credentials);

            ExchangeMessageInfoCollection list = client.ListMessages(client.MailboxInfo.DeletedItemsUri);
            Console.WriteLine(list[0].MessageInfoType.ToString());
            //ExEnd: GetMessageTypeFromExchangeMessageInfo
        }
    }
}
