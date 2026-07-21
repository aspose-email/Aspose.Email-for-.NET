// Demonstrates how to give an Outlook task a daily recurrence that never ends.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class SetDailyNeverEndRecurrence
    {
        public static void Run()
        {
            var startDate = new DateTime(2015, 7, 16);
            var dueDate = new DateTime(2015, 7, 16);

            var task = new MapiTask("This is test task", "Sample Body", startDate, dueDate)
            {
                State = MapiTaskState.NotAssigned
            };

            // NeverEnd means the pattern has neither an end date nor an occurrence
            // limit, so no OccurrenceCount or EndDate is needed here.
            task.Recurrence = new MapiCalendarDailyRecurrencePattern
            {
                PatternType = MapiCalendarRecurrencePatternType.Day,
                Period = 1,
                EndType = MapiCalendarRecurrenceEndType.NeverEnd
            };

            var outputPath = Data.Out/"SetDailyNeverEndRecurrence_out.msg";
            task.Save(outputPath, TaskSaveFormat.Msg);

            Console.WriteLine("Task recurs every day and never ends.");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}
