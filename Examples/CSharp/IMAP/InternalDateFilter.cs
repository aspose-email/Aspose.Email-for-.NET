using System;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Tools.Search;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.IMAP
{
    class InternalDateFilter
    {
        public static void Run()
        {
            // ExStart:InternalDateFilter
            // Connect and log in to IMAP
            const string host = "host";
            const int port = 143;
            const string username = "user@host.com";
            const string password = "password";
            ImapClient client = new ImapClient(host, port, username, password);
            client.SelectFolder("Inbox");

            // Set conditions, Subject contains "Newsletter", Emails that arrived today
            ImapQueryBuilder builder = new ImapQueryBuilder();
            builder.Subject.Contains("Newsletter");
            builder.InternalDate.On(DateTime.Now);

            // Build the query and Get list of messages
            MailQuery query = builder.GetQuery();
            ImapMessageInfoCollection messages = client.ListMessages(query);
            foreach (ImapMessageInfo info in messages)
            {
                Console.WriteLine("Internal Date: " + info.InternalDate);
            }
            
            // Disconnect from IMAP
            client.Dispose();
            // ExEnd:InternalDateFilter
        }
    }
}