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
    class ListMessagesByID
    {
        public static void Run()
        {
            try
            {

                // ExStart:ListMessagesByID
                // Create instance of ExchangeClient class by giving credentials
                ExchangeClient client = new ExchangeClient("https://MachineName/exchange/Username", "username", "password", "domain");

                // Call ListMessages method to list messages info from Inbox
                ExchangeMessageInfoCollection msgCollection = client.ListMessagesById(client.MailboxInfo.InboxUri, "23A747F0C7A5DB4BAB299C2BE2383FD556E630DD@machinename.local");

                // Loop through the collection to display the basic information
                foreach (ExchangeMessageInfo msgInfo in msgCollection)
                {
                    Console.WriteLine("Subject: " + msgInfo.Subject);
                    Console.WriteLine("From: " + msgInfo.From.ToString());
                    Console.WriteLine("To: " + msgInfo.To.ToString());
                    Console.WriteLine("Message ID: " + msgInfo.MessageId);
                    Console.WriteLine("Unique URI: " + msgInfo.UniqueUri);
                    Console.WriteLine("==================================");
                }
                // ExEnd:ListMessagesByID
            }
            catch (Exception ex)
            {

                Console.Write(ex.Message);
            }
        }
    }
}