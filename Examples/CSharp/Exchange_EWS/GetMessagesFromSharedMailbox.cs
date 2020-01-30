using System;
using System.Net;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Clients.Exchange;
using Aspose.Email.Mapi;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from https://www.nuget.org/packages/Aspose.Email/, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using https://forum.aspose.com/c/email
*/

namespace Aspose.Email.Examples.CSharp.Email.Exchange_EWS
{
    public class GetMessagesFromSharedMailbox
    {
        public static void Run()
        {
            //ExStart: 1
            const string mailboxUri = "<HOST>";
            const string domain = "";
            const string username = "<EMAIL ADDRESS>";
            const string password = "<PASSWORD>";
            const string sharedEmail = "<SHARED EMAIL ADDRESS>";
            NetworkCredential credentials = new NetworkCredential(username, password, domain);
            IEWSClient client = EWSClient.GetEWSClient(mailboxUri, credentials);

            string[] items = client.ListItems(sharedEmail, "Inbox");

            foreach (string item in items)
            {
                MapiMessage msg = client.FetchItem(item);
                Console.WriteLine("Subject:" + msg.Subject);
            }
            client.Dispose();
            //ExEnd: 1
            Console.WriteLine("GetMessagesFromSharedMailbox executed successfully");
        }
    }
}