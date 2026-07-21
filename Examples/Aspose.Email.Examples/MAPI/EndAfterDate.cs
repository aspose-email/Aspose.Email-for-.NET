// Demonstrates how to create a recurring MAPI task that ends after a specific date.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class EndAfterDate
    {
        public static void Run()
        {
            // Outlook stores task dates in UTC, so shift the local dates by the local
            // offset to keep the task on the intended day.
            var offset = TimeZoneInfo.Local.GetUtcOffset(DateTime.Now);
            var startDate = new DateTime(2015, 7, 1).Add(offset);
            var dueDate = new DateTime(2015, 7, 1).Add(offset);
            var endByDate = new DateTime(2018, 7, 1).Add(offset);

            var task = new MapiTask("This is test task", "Sample Body", startDate, dueDate)
            {
                State = MapiTaskState.NotAssigned
            };

            // EndAfterDate is what makes EndDate the stopping condition; the pattern
            // itself repeats on day 15 every 12 months.
            var recurrence = new MapiCalendarMonthlyRecurrencePattern
            {
                Day = 15,
                Period = 12,
                PatternType = MapiCalendarRecurrencePatternType.Month,
                EndType = MapiCalendarRecurrenceEndType.EndAfterDate,
                EndDate = endByDate,
                OccurrenceCount = 3
            };

            task.Recurrence = recurrence;

            var outputPath = Data.Out/"EndAfterDate_out.msg";
            task.Save(outputPath, TaskSaveFormat.Msg);

            Console.WriteLine($"Task recurs on day {recurrence.Day} every {recurrence.Period} months.");
            Console.WriteLine($"Ends on: {endByDate:d}");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}
