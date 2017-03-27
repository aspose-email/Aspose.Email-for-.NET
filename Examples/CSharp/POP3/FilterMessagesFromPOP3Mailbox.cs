using System;
using Aspose.Email.Clients.Pop3;
using Aspose.Email.Tools.Search;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.POP3
{
    class FilterMessagesFromPOP3Mailbox
    {
        public static void Run()
        {
            // ExStart:FilterMessagesFromPOP3Mailbox
            // Connect and log in to POP3
            const string host = "host";
            const int port = 110;
            const string username = "user@host.com";
            const string password = "password";
            Pop3Client client = new Pop3Client(host, port, username, password);

            // Set conditions, Subject contains "Newsletter", Emails that arrived today
            MailQueryBuilder builder = new MailQueryBuilder();      
            builder.Subject.Contains("Newsletter");
            builder.InternalDate.On(DateTime.Now);
            // Build the query and Get list of messages
            MailQuery query = builder.GetQuery();
            Pop3MessageInfoCollection messages = client.ListMessages(query);
            Console.WriteLine("Pop3: " + messages.Count + " message(s) found.");
            // ExEnd:FilterMessagesFromPOP3Mailbox
        }
    }
}