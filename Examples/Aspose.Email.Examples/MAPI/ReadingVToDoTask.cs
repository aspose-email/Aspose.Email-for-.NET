// Demonstrates how to read a task from an iCalendar VTODO file and save it as an
// Outlook MSG task.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class ReadingVToDoTask
    {
        public static void Run()
        {
            var task = MapiTask.FromVTodo(Data.Mapi/"VtodoTask.ics");

            var outputPath = Data.Out/"VToDo_out.msg";
            task.Save(outputPath, TaskSaveFormat.Msg);

            Console.WriteLine($"Subject:  {task.Subject}");
            Console.WriteLine($"Status:   {task.Status}");
            Console.WriteLine($"Due date: {task.DueDate:d}");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}
