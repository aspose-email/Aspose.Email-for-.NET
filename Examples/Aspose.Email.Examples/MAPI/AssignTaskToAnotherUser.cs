// Demonstrates the delegation side of a MAPI task: who assigned it, who it went to,
// and where it sits in the assignment handshake.
//
// MapiTaskUsers carries the people. Mode / State / AcceptanceState track the handshake
// itself, and they are driven by actually sending and answering the assignment through
// a mail client - filling in the user fields below does not move them on its own, so
// they stay at NotAssigned here.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class AssignTaskToAnotherUser
    {
        public static void Run()
        {
            var task = new MapiTask("Quarterly report", "Please prepare the Q3 numbers.",
                DateTime.Now, DateTime.Now.AddDays(7));

            Console.WriteLine($"Task: {task.Subject}, due {task.DueDate:yyyy-MM-dd}");

            // Who handed the task over, and to whom.
            task.Users.Assigner = new MapiElectronicAddress
            {
                EmailAddress = "manager@example.com",
                DisplayName = "Morgan Manager"
            };
            task.Users.Owner = "Alex Analyst";
            task.Users.LastUser = "Morgan Manager";
            task.Users.LastAssigner = "Morgan Manager";
            task.Users.Delegator = "Morgan Manager";
            task.Users.Ownership = MapiTaskOwnership.AssignersCopy;

            // Attendees are the recipients the assignment is sent to.
            task.Users.Attendees.Add("analyst@example.com", "SMTP", "Alex Analyst", MapiRecipientType.MAPI_TO);

            task.Priority = MapiTaskPriority.High;
            task.Status = MapiTaskStatus.InProgress;
            task.PercentComplete = 25;

            Console.WriteLine("\nPeople:");
            Console.WriteLine($"  assigner:   {task.Users.Assigner.DisplayName} <{task.Users.Assigner.EmailAddress}>");
            Console.WriteLine($"  owner:      {task.Users.Owner}");
            Console.WriteLine($"  delegator:  {task.Users.Delegator}");
            Console.WriteLine($"  ownership:  {task.Users.Ownership}");
            Console.WriteLine($"  attendees:  {task.Users.Attendees.Count}");

            // Unchanged by everything above - only a real send moves these on.
            Console.WriteLine("\nHandshake state:");
            Console.WriteLine($"  mode:       {task.Mode}");
            Console.WriteLine($"  state:      {task.State}");
            Console.WriteLine($"  acceptance: {task.AcceptanceState}");
            Console.WriteLine($"  flags:      {task.Flags}");

            var outputPath = Data.Out/"AssignTaskToAnotherUser_out.msg";
            task.Save(outputPath, TaskSaveFormat.Msg);

            var reloaded = (MapiTask)MapiMessage.Load(outputPath).ToMapiMessageItem();

            Console.WriteLine("\nAfter a round trip through MSG:");
            Console.WriteLine($"  assigner:   {reloaded.Users.Assigner.EmailAddress}");
            Console.WriteLine($"  last user:  {reloaded.Users.LastUser}");
            Console.WriteLine($"  attendees:  {reloaded.Users.Attendees.Count}");
            Console.WriteLine($"  status:     {reloaded.Status}, {reloaded.PercentComplete}% complete");

            // Priority is not among the properties the task MSG format keeps, so it
            // comes back at its default rather than the value set above.
            Console.WriteLine($"  priority:   set to {task.Priority}, reloaded as {reloaded.Priority}");

            Console.WriteLine($"\nSaved to {outputPath}");
        }
    }
}
