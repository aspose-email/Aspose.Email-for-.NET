using System;
using Aspose.Email.Exchange;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Exchange
{
    class GetMailboxInformationFromExchangeServer
    {
        public static void Run()
        {
            // The path to the File directory.
            // ExStart:GetMailboxInformationFromExchangeServer
 
            // Create instance of ExchangeClient class by giving credentials
            ExchangeClient client = new ExchangeClient("http:// MachineName/exchange/Username","Username", "password", "domain");

            // Get mailbox size
            Console.WriteLine("Mailbox size: " + client.GetMailboxSize() + " bytes");

            // Get exchange mailbox info
            ExchangeMailboxInfo mailboxInfo = client.GetMailboxInfo();

            // Get Mailbox URI
            Console.WriteLine("Mailbox URI: " + mailboxInfo.MailboxUri);

            // Get Inbox folder URI
            Console.WriteLine("Inbox folder URI: " + mailboxInfo.InboxUri);

            // Get Sent Items folder URI
            Console.WriteLine("Sent Items URI: " + mailboxInfo.SentItemsUri);

            // Get Drafts folder URI
            Console.WriteLine("Drafts folder URI: " + mailboxInfo.DraftsUri);
            // ExEnd:GetMailboxInformationFromExchangeServer
        }
    }
}