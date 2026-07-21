// Demonstrates how to create a detailed MAPI task with properties and save it as MSG.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class CreatingAndSavingOutlookTasks
    {
        public static void Run()
        {
            var task = new MapiTask("To Do", "Just click and type to add new task", DateTime.Now, DateTime.Now.AddDays(3));
            task.PercentComplete = 20;
            task.EstimatedEffort = 5;
            task.ActualEffort = 20;
            task.History = MapiTaskHistory.Assigned;
            task.LastUpdate = DateTime.Now;
            task.Users.Owner = "Darius";
            task.Users.LastAssigner = "Harkness";
            task.Users.LastDelegate = "Harkness";
            task.Users.Ownership = MapiTaskOwnership.AssignersCopy;
            task.Users.Delegator = "Test Delegator";
            task.Companies = new string[] { "company1", "company2", "company3" };
            task.Categories = new string[] { "category1", "category2", "category3" };
            task.Mileage = "Some test mileage";
            task.Billing = "Test billing information";
            task.Status = MapiTaskStatus.Complete;

            var outputPath = Data.Out/"MapiTask.msg";
            task.Save(outputPath, TaskSaveFormat.Msg);

            Console.WriteLine($"Task:     {task.Subject}");
            Console.WriteLine($"Status:   {task.Status}, {task.PercentComplete}% complete");
            Console.WriteLine($"Owner:    {task.Users.Owner}, delegated by {task.Users.Delegator}");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}
