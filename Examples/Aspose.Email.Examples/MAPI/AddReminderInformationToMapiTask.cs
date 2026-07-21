// Demonstrates how to create a MAPI task with a reminder set to a specific time and audio file.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class AddReminderInformationToMapiTask
    {
        public static void Run()
        {
            var task = new MapiTask("task with reminder", "this is a body", DateTime.Now, DateTime.Now.AddHours(1))
            {
                // ReminderSet is what switches the reminder on; the file parameter makes
                // Outlook play a sound rather than only showing the pop-up.
                ReminderSet = true,
                ReminderTime = DateTime.Now,
                ReminderFileParameter = Data.Mapi/"Alarm01.wav"
            };

            var outputPath = Data.Out/"AddReminderInformationToMapiTask_out.msg";
            task.Save(outputPath, TaskSaveFormat.Msg);

            Console.WriteLine($"Task:     {task.Subject}");
            Console.WriteLine($"Reminder: {task.ReminderTime:g}, sound {task.ReminderFileParameter}");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}
