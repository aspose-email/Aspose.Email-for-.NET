// Demonstrates how to give an Outlook task a daily recurrence that stops after a
// fixed number of occurrences.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class SetDailyOccurrenceCount
    {
        public static void Run()
        {
            var startDate = new DateTime(2015, 7, 16);
            var dueDate = new DateTime(2015, 7, 16);

            var task = new MapiTask("This is test task", "Sample Body", startDate, dueDate)
            {
                State = MapiTaskState.NotAssigned
            };

            // EndAfterNOccurrences makes OccurrenceCount the limit: the task repeats
            // five times and then stops.
            var recurrence = new MapiCalendarDailyRecurrencePattern
            {
                PatternType = MapiCalendarRecurrencePatternType.Day,
                Period = 1,
                WeekStartDay = DayOfWeek.Sunday,
                EndType = MapiCalendarRecurrenceEndType.EndAfterNOccurrences,
                OccurrenceCount = 5
            };

            task.Recurrence = recurrence;

            var outputPath = Data.Out/"SetDailyOccurrenceCount_out.msg";
            task.Save(outputPath, TaskSaveFormat.Msg);

            Console.WriteLine($"Task recurs every day, {recurrence.OccurrenceCount} times.");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}
