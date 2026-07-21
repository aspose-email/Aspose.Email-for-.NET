// Demonstrates how to create a MAPI task and add it to a PST tasks folder.

using System;
using System.IO;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.MAPI
{
    internal static class AddMapiTaskToPst
    {
        public static void Run()
        {
            var task = new MapiTask("To Do", "Just click and type to add new task", DateTime.Now, DateTime.Now.AddDays(3));
            task.PercentComplete = 20;
            task.EstimatedEffort = 2000;
            task.ActualEffort = 20;
            task.History = MapiTaskHistory.Assigned;
            task.LastUpdate = DateTime.Now;
            task.Users.Owner = "Darius";
            task.Users.LastAssigner = "Harkness";
            task.Users.LastDelegate = "Harkness";
            task.Users.Ownership = MapiTaskOwnership.AssignersCopy;

            var path = Data.Out/"AddMapiTaskToPST_out.pst";

            if (File.Exists(path))
                File.Delete(path);

            using (var pst = PersonalStorage.Create(path, FileFormatVersion.Unicode))
            {
                var taskFolder = pst.CreatePredefinedFolder("Tasks", StandardIpmFolder.Tasks);
                taskFolder.AddMapiMessageItem(task);

                Console.WriteLine($"Task:     {task.Subject}");
                Console.WriteLine($"Complete: {task.PercentComplete}%");
                Console.WriteLine($"Owner:    {task.Users.Owner}");
                Console.WriteLine($"Saved to {path}");
            }
        }
    }
}
