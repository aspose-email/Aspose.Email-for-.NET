using Aspose.Email.Clients.Exchange;
using Aspose.Email.Clients.Exchange.WebService;
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
    class FilterTasksFromServer
    {
        public static void Run()
        {
            // ExStart:FilterTasksFromServer
            // Set mailboxURI, Username, password, domain information
            string mailboxUri = "https://ex2010/ews/exchange.asmx";
            string username = "test.exchange";
            string password = "pwd";
            string domain = "ex2010.local";
            NetworkCredential credentials = new NetworkCredential(username, password, domain);
            IEWSClient client = EWSClient.GetEWSClient(mailboxUri, credentials);

            ExchangeQueryBuilder queryBuilder = null;
            MailQuery query = null;
            ExchangeTask fetchedTask = null;
            ExchangeMessageInfoCollection messageInfoCol = null;
            client.TimezoneId = "Central Europe Standard Time";
            Array values = Enum.GetValues(typeof(ExchangeTaskStatus));

            //Now retrieve the tasks with specific statuses
            foreach (ExchangeTaskStatus status in values)
            {
                queryBuilder = new ExchangeQueryBuilder();
                queryBuilder.TaskStatus.Equals(status);
                query = queryBuilder.GetQuery();
                messageInfoCol = client.ListMessages(client.MailboxInfo.TasksUri, query);
                fetchedTask = client.FetchTask(messageInfoCol[0].UniqueUri);
            }

            //retrieve all other than specified
            foreach (ExchangeTaskStatus status in values)
            {
                queryBuilder = new ExchangeQueryBuilder();
                queryBuilder.TaskStatus.NotEquals(status);
                query = queryBuilder.GetQuery();
                messageInfoCol = client.ListMessages(client.MailboxInfo.TasksUri, query);
            }

            //specifying multiple criterion
            ExchangeTaskStatus[] selectedStatuses = new ExchangeTaskStatus[]
                    {
            ExchangeTaskStatus.Completed,
            ExchangeTaskStatus.InProgress
                    };
            queryBuilder = new ExchangeQueryBuilder();
            queryBuilder.TaskStatus.In(selectedStatuses);
            query = queryBuilder.GetQuery();
            messageInfoCol = client.ListMessages(client.MailboxInfo.TasksUri, query);

            queryBuilder = new ExchangeQueryBuilder();
            queryBuilder.TaskStatus.NotIn(selectedStatuses);
            query = queryBuilder.GetQuery();
            messageInfoCol = client.ListMessages(client.MailboxInfo.TasksUri, query);
            //ExEnd:FilterTasksFromServer
        }
    }
}
