// Demonstrates the ways to filter tasks by status.
//
// ExchangeQueryBuilder.TaskStatus supports Equals, NotEquals, In and NotIn, so tasks can
// be selected by one status, all but one, or any of several. The example counts the
// tasks of the Tasks folder for each kind of condition.

using System;
using Aspose.Email.Clients.Exchange;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Tools.Search;

namespace Aspose.Email.Examples.EWS
{
    internal static class FilterTasksFromServer
    {
        public static void Run()
        {
            var openStatuses = new[] { ExchangeTaskStatus.NotStarted, ExchangeTaskStatus.InProgress };

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                foreach (ExchangeTaskStatus status in Enum.GetValues(typeof(ExchangeTaskStatus)))
                {
                    var builder = new ExchangeQueryBuilder();
                    builder.TaskStatus.Equals(status);
                    Count(client, $"Status is {status}", builder.GetQuery());
                }

                var notCompleted = new ExchangeQueryBuilder();
                notCompleted.TaskStatus.NotEquals(ExchangeTaskStatus.Completed);
                Count(client, "Status is not Completed", notCompleted.GetQuery());

                var open = new ExchangeQueryBuilder();
                open.TaskStatus.In(openStatuses);
                Count(client, "Not started or in progress", open.GetQuery());

                var closed = new ExchangeQueryBuilder();
                closed.TaskStatus.NotIn(openStatuses);
                Count(client, "Neither not started nor in progress", closed.GetQuery());
            }
        }

        private static void Count(IEWSClient client, string description, MailQuery query)
        {
            var tasks = client.ListTasks(client.MailboxInfo.TasksUri, query);
            Console.WriteLine($"{description,-38} {tasks.Count} task(s)");
        }
    }
}
