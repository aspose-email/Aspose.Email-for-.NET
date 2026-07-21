// Demonstrates how to give an Outlook task a yearly recurrence that stops after a
// fixed number of occurrences.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class YearlyEndAfterNoccurrences
    {
        public static void Run()
        {
            // Outlook stores task dates in UTC, so shift the local dates by the local
            // offset to keep the task on the intended day.
            var offset = TimeZoneInfo.Local.GetUtcOffset(DateTime.Now);

            var startDate = new DateTime(2015, 7, 1).Add(offset);
            var dueDate = new DateTime(2015, 7, 1).Add(offset);

            var task = new MapiTask("This is test task", "Sample Body", startDate, dueDate)
            {
                State = MapiTaskState.NotAssigned
            };

            // A yearly pattern is a monthly one with a period of 12 months; the task
            // repeats three times and then stops.
            var recurrence = new MapiCalendarMonthlyRecurrencePattern
            {
                Day = 15,
                Period = 12,
                PatternType = MapiCalendarRecurrencePatternType.Month,
                EndType = MapiCalendarRecurrenceEndType.EndAfterNOccurrences,
                OccurrenceCount = 3
            };

            task.Recurrence = recurrence;

            var outputPath = Data.Out/"YearlyEndAfterNoccurrences_out.msg";
            task.Save(outputPath, TaskSaveFormat.Msg);

            Console.WriteLine($"Task recurs yearly, {recurrence.OccurrenceCount} occurrence(s).");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}
