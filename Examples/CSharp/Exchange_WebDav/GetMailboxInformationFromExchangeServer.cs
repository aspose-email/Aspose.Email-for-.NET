using System;
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
    class GetMailboxInformationFromExchangeServer
    {
        public static void Run()
        {
            try
            {
                // ExStart:GetMailboxInformationFromExchangeServer 
                // Create instance of ExchangeClient class by giving credentials
                ExchangeClient client = new ExchangeClient("https://MachineName/exchange/Username", "Username", "password", "domain");
                
                // Get mailbox size, exchange mailbox info, Mailbox, Inbox folder, Sent Items folder URI , Drafts folder URI
                Console.WriteLine("Mailbox size: " + client.GetMailboxSize() + " bytes");
                ExchangeMailboxInfo mailboxInfo = client.GetMailboxInfo();
                Console.WriteLine("Mailbox URI: " + mailboxInfo.MailboxUri);
                Console.WriteLine("Inbox folder URI: " + mailboxInfo.InboxUri);
                Console.WriteLine("Sent Items URI: " + mailboxInfo.SentItemsUri);
                Console.WriteLine("Drafts folder URI: " + mailboxInfo.DraftsUri);
                // ExEnd:GetMailboxInformationFromExchangeServer
            }
            catch (Exception ex)
            {

                Console.Write(ex.Message);
            }
        }
    }
}