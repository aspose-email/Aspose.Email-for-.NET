using Aspose.Email.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Email
{
    class Program
    {
        static void Main(string[] args)
        {
            DateTime startDate = new DateTime(2015, 04, 30, 10, 00, 00);
            MapiTask task = new MapiTask("abc", "def", startDate, startDate.AddHours(1));
            task.State = MapiTaskState.NotAssigned;

            // Set the weekly recurrence
            var rec = new MapiCalendarDailyRecurrencePattern
            {
                PatternType = MapiCalendarRecurrencePatternType.Day,
                Period = 1,
                WeekStartDay = DayOfWeek.Sunday,
                EndType = MapiCalendarRecurrenceEndType.NeverEnd,
                OccurrenceCount = 0,
            };
            task.Recurrence = rec;
            task.Save("AsposeDaily.msg", TaskSaveFormat.Msg);
        }
    }
}
