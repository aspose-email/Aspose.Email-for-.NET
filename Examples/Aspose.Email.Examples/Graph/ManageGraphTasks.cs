// Demonstrates Microsoft To Do through Graph: the task lists of a mailbox, and the
// tasks inside them.

using System;
using System.Linq;
using Aspose.Email.Clients.Graph;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.Graph
{
    internal static class ManageGraphTasks
    {
        public static void Run()
        {
            if (!ClientBuilder.IsGraphConfigured)
            {
                GraphExampleInfo.PrintNotConfigured();
                return;
            }

            using (var client = ClientBuilder.Graph(AuthType.ModernWithAppPermission))
            {
                var taskLists = client.ListTaskLists(null);
                Console.WriteLine($"{taskLists.Count} task list(s):");

                foreach (var list in taskLists)
                {
                    Console.WriteLine($"  {list.DisplayName} (well-known name: {list.WellknownName}, " +
                                      $"owner: {list.IsOwner}, shared: {list.IsShared})");
                }

                // A list of your own, alongside the built-in ones.
                var created = client.CreateTaskList(new TaskListInfo { DisplayName = "Aspose.Email examples" });
                Console.WriteLine($"\nCreated list: {created.DisplayName} ({created.ItemId})");

                var task = new MapiTask
                {
                    Subject = "Prepare the quarterly report",
                    DueDate = DateTime.UtcNow.Date.AddDays(7),
                    Status = MapiTaskStatus.NotStarted
                };

                var createdTask = client.CreateTask(task, created.ItemId);
                Console.WriteLine($"Created task: {createdTask.Subject}, due {createdTask.DueDate:yyyy-MM-dd}");

                var tasks = client.ListTasks(created.ItemId, null);
                Console.WriteLine($"{tasks.Count} task(s) in the list");

                var fetched = client.FetchTask(createdTask.ItemId);
                Console.WriteLine($"Fetched: {fetched.Subject}, status {fetched.Status}");

                fetched.Status = MapiTaskStatus.InProgress;
                fetched.PercentComplete = 30;

                var updated = client.UpdateTask(fetched, new UpdateSettings { SkipAttachments = true });
                Console.WriteLine($"Updated: {updated.Status}, {updated.PercentComplete}% complete");

                // Renaming the list, then reading it back by id.
                created.DisplayName = "Aspose.Email examples (done)";
                var renamed = client.UpdateTaskList(created);
                Console.WriteLine($"Renamed list to: {renamed.DisplayName}");

                var reloaded = client.GetTaskList(renamed.ItemId);
                Console.WriteLine($"Reloaded list: {reloaded.DisplayName}");

                client.DeleteTaskList(renamed.ItemId);
                Console.WriteLine("Deleted the task list.");
            }
        }
    }
}
