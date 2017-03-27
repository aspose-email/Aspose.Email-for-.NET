using Aspose.Email.Clients.Pop3;
using Aspose.Email.Tools.Search;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.POP3
{
    class GetMessagesUsingSpecificCriteria
    {
        public static void Run()
        {
            // Connect and log in to POP3
            const string host = "host";
            const int port = 110;
            const string username = "user@host.com";
            const string password = "password";
            Pop3Client client = new Pop3Client(host, port, username, password);

            try
            {

                // ExStart:GetEmailsWithTodayDate
                // Emails that arrived today
                MailQueryBuilder builder = new MailQueryBuilder();
                builder.InternalDate.On(DateTime.Now);
                // ExEnd:GetEmailsWithTodayDate
                
                // Build the query and Get list of messages
                MailQuery query = builder.GetQuery();
                Pop3MessageInfoCollection messages = client.ListMessages(query);
                Console.WriteLine("Pop3: " + messages.Count + " message(s) found.");

                builder = new MailQueryBuilder();

                // ExStart:GetEmailsOverDateRange
                // Emails that arrived in last 7 days
                builder.InternalDate.Before(DateTime.Now);
                builder.InternalDate.Since(DateTime.Now.AddDays(-7));
                // ExEnd:GetEmailsOverDateRange

                // Build the query and Get list of messages
                query = builder.GetQuery();
                messages = client.ListMessages(query);
                Console.WriteLine("Pop3: " + messages.Count + " message(s) found.");

                builder = new MailQueryBuilder();

                // ExStart:GetSpecificSenderEmails
                // Get emails from specific sender
                builder.From.Contains("saqib.razzaq@127.0.0.1");
                // ExEnd:GetSpecificSenderEmails

                // Build the query and Get list of messages
                query = builder.GetQuery();
                messages = client.ListMessages(query);
                Console.WriteLine("Pop3: " + messages.Count + " message(s) found.");

                builder = new MailQueryBuilder();

                // ExStart:GetSpecificDomainEmails
                // Get emails from specific domain
                builder.From.Contains("SpecificHost.com");
                // ExEnd:GetSpecificDomainEmails

                // Build the query and Get list of messages
                query = builder.GetQuery();
                messages = client.ListMessages(query);
                Console.WriteLine("Pop3: " + messages.Count + " message(s) found.");

                builder = new MailQueryBuilder();

                // ExStart:GetSpecificRecipientEmails
                // Get emails sent to specific recipient
                builder.To.Contains("recipient");
                // ExEnd:GetSpecificRecipientEmails

                // Build the query and Get list of messages
                query = builder.GetQuery();
                messages = client.ListMessages(query);
                Console.WriteLine("Pop3: " + messages.Count + " message(s) found.");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
