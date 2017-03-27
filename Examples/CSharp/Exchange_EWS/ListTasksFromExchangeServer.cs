using Aspose.Email.Calendar;
using Aspose.Email.Clients.Exchange;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Mime;
using Aspose.Email.Tools.Search;
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
    class ListTasksFromExchangeServer
    {
        public static void Run()
        {
            // ExStart:ListTasksFromExchangeServerWithEWS
            // Set mailboxURI, Username, password, domain information
            string mailboxUri = "https://ex2010/ews/exchange.asmx";
            string username = "test.exchange";
            string password = "pwd";
            string domain = "ex2010.local";
            NetworkCredential credentials = new NetworkCredential(username, password, domain);
            IEWSClient client = EWSClient.GetEWSClient(mailboxUri, credentials);

            //Listing Tasks from Server
            client.TimezoneId = "Central Europe Standard Time";
            TaskCollection taskCollection = client.ListTasks(client.MailboxInfo.TasksUri);

            //print retrieved tasks' details
            foreach (ExchangeTask task in taskCollection)
            {
                Console.WriteLine(task.TimezoneId);
                Console.WriteLine(task.Subject);
                Console.WriteLine(task.StartDate);
                Console.WriteLine(task.DueDate);
            }

            //Listing Tasks from server based on Query - Completed and In-Progress
            ExchangeQueryBuilder builder = new ExchangeQueryBuilder();
            ExchangeTaskStatus[] selectedStatuses = new ExchangeTaskStatus[]
            {
                ExchangeTaskStatus.Completed,
                ExchangeTaskStatus.InProgress
            };
            builder.TaskStatus.In(selectedStatuses);

            MailQuery query = builder.GetQuery();

            taskCollection = client.ListTasks(client.MailboxInfo.TasksUri, query);

            //print retrieved tasks' details
            foreach (ExchangeTask task in taskCollection)
            {
                Console.WriteLine(task.TimezoneId);
                Console.WriteLine(task.Subject);
                Console.WriteLine(task.StartDate);
                Console.WriteLine(task.DueDate);
            }
            //ExEnd:ListTasksFromExchangeServerWithEWS
        }
    }
}
