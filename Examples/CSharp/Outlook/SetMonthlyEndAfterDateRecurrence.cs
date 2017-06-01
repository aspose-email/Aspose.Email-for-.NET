using Aspose.Email.Mapi;
using Aspose.Email.Calendar.Recurrences;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

/* This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET 
   API reference when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq 
   for more information. If you do not wish to use NuGet, you can manually download 
   Aspose.Email for .NET API from http://www.aspose.com/downloads, 
   install it and then add its reference to this project. For any issues, questions or suggestions 
   please feel free to contact us using http://www.aspose.com/community/forums/default.aspx            
*/

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class SetMonthlyEndAfterDateRecurrence
    {
        public static void Run()
        {
            // ExStart:SetMonthlyEndAfterDateRecurrence
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Outlook();

            TimeZone localZone = TimeZone.CurrentTimeZone;
            TimeSpan ts = localZone.GetUtcOffset(DateTime.Now);
            DateTime StartDate = new DateTime(2015, 7, 1);
            StartDate = StartDate.Add(ts);

            DateTime DueDate = new DateTime(2015, 7, 1);
            DateTime endByDate = new DateTime(2015, 12, 31);
            DueDate = DueDate.Add(ts);
            endByDate = endByDate.Add(ts);

            MapiTask task = new MapiTask("This is test task", "Sample Body", StartDate, DueDate);
            task.State = MapiTaskState.NotAssigned;

            // Set the Monthly recurrence
            var recurrence = new MapiCalendarMonthlyRecurrencePattern
            {
                Day = 15,
                Period = 1,
                PatternType = MapiCalendarRecurrencePatternType.Month,
                EndType = MapiCalendarRecurrenceEndType.EndAfterDate,
                OccurrenceCount = GetOccurrenceCount(StartDate, endByDate, "FREQ=MONTHLY;BYMONTHDAY=15;INTERVAL=1"),
                WeekStartDay = DayOfWeek.Monday,
                EndDate = endByDate
            };

            task.Recurrence = recurrence;
            //task.Save(dataDir + "SetMonthlyEndAfterDateRecurrence_out.msg", TaskSaveFormat.Msg);
            // ExEnd:SetMonthlyEndAfterDateRecurrence
        }

        private static uint GetOccurrenceCount(DateTime start, DateTime endBy, string rrule)
        {
            CalendarRecurrence pattern = new CalendarRecurrence(string.Format("DTSTART:{0}\r\nRRULE:{1}", start.ToString("yyyyMMdd"), rrule));
            DateCollection dates = pattern.GenerateOccurrences(start, endBy);
            return (uint)dates.Count;
        }
    }
}
