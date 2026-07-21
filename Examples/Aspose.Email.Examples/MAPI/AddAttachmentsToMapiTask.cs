// Demonstrates how to create a MAPI task with a text attachment and save it as MSG.

using System;
using System.Text;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class AddAttachmentsToMapiTask
    {
        public static void Run()
        {
            var task = new MapiTask(
                "Task with attachment",
                "Test body of task with attachment",
                DateTime.Now,
                DateTime.Now.AddHours(1));

            // An attachment can be added straight from a byte array, without a file on disk.
            task.Attachments.Add("Test attachment name", Encoding.Unicode.GetBytes("Test attachment body"));

            var outputPath = Data.Out/"AddAttachmentsToMapiTask_out.msg";
            task.Save(outputPath, TaskSaveFormat.Msg);

            Console.WriteLine($"Task: {task.Subject} ({task.Attachments.Count} attachment(s))");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}
