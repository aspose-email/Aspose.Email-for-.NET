// Demonstrates listing tasks, all or filtered by status.
//
// ListTasks returns ExchangeTask objects with subject, status and dates. The overload
// with a query narrows the list - here to tasks that are completed or in progress.

using System;
using Aspose.Email.Clients.Exchange;
using Aspose.Email.Clients.Exchange.WebService;

namespace Aspose.Email.Examples.EWS
{
    internal static class ListTasksFromExchangeServer
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var tasksUri = client.MailboxInfo.TasksUri;

                var all = client.ListTasks(tasksUri);
                Console.WriteLine($"{all.Count} task(s):");
                foreach (ExchangeTask task in all)
                    {
                        Console.WriteLine($"  {task.Subject}: {task.Status}, " +
                                          $"start {task.StartDate:d}, due {task.DueDate:d}");
                    }

                var builder = new ExchangeQueryBuilder();
                builder.TaskStatus.In(new[] { ExchangeTaskStatus.Completed, ExchangeTaskStatus.InProgress });

                var active = client.ListTasks(tasksUri, builder.GetQuery());
                Console.WriteLine($"\nCompleted or in progress: {active.Count}");
            }
        }
    }
}
