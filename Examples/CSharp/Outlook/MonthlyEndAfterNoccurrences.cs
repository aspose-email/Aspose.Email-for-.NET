using System;
using Aspose.Email.Mapi;
using Aspose.Email.Calendar.Recurrences;

/* This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET 
   API reference when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq 
   for more information. If you do not wish to use NuGet, you can manually download 
   Aspose.Email for .NET API from http://www.aspose.com/downloads, 
   install it and then add its reference to this project. For any issues, questions or suggestions 
   please feel free to contact us using http://www.aspose.com/community/forums/default.aspx            
*/

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class MonthlyEndAfterNoccurrences
    {
        public static void Run()
        {
            // ExStart:MonthlyEndAfterNoccurrences
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Outlook();
            TimeZone localZone = TimeZone.CurrentTimeZone;
            TimeSpan ts = localZone.GetUtcOffset(DateTime.Now);
            DateTime StartDate = new DateTime(2015, 7, 16);
            StartDate = StartDate.Add(ts);

            DateTime DueDate = new DateTime(2015, 7, 16);
            DateTime endByDate = new DateTime(2015, 12, 31);
            DueDate = DueDate.Add(ts);
            endByDate = endByDate.Add(ts);

            MapiTask task = new MapiTask("This is test task", "Sample Body", StartDate, DueDate);
            task.State = MapiTaskState.NotAssigned;

            // Set the Monthly recurrence
            var rec = new MapiCalendarMonthlyRecurrencePattern
            {
                Day = 15,
                Period = 1,
                PatternType = MapiCalendarRecurrencePatternType.Month,
                EndType = MapiCalendarRecurrenceEndType.EndAfterNOccurrences,
                OccurrenceCount = GetOccurrenceCount(StartDate, endByDate, "FREQ=MONTHLY;BYMONTHDAY=15;INTERVAL=1"),
                WeekStartDay = DayOfWeek.Monday,

            };

            if (rec.OccurrenceCount == 0)
            {
                rec.OccurrenceCount = 1;
            }

            task.Recurrence = rec;
            //task.Save(dataDir + "Monthly_out.msg", TaskSaveFormat.Msg);
            // ExEnd:MonthlyEndAfterNoccurrences

            // ExStart:SetFixNumberOfOccurrences
            // Set the Monthly recurrence
            var records = new MapiCalendarMonthlyRecurrencePattern
            {
                Day = 15,
                Period = 1,
                PatternType = MapiCalendarRecurrencePatternType.Month,
                EndType = MapiCalendarRecurrenceEndType.EndAfterNOccurrences,
                OccurrenceCount = 5,
                WeekStartDay = DayOfWeek.Monday
            };
            // ExEnd:SetFixNumberOfOccurrences

            task.Recurrence = records;
            //task.Save(dataDir + "SetFixNumberOfOccurrences_out.msg", TaskSaveFormat.Msg);
        }

        // ExStart:EventsBetweenTheTwoDates     
        private static uint GetOccurrenceCount(DateTime start, DateTime endBy, string rrule)
        {
            CalendarRecurrence pattern = new CalendarRecurrence(string.Format("DTSTART:{0}\r\nRRULE:{1}", start.ToString("yyyyMMdd"), rrule));
            DateCollection dates = pattern.GenerateOccurrences(start, endBy);       
            return (uint)dates.Count;
        }
        // ExEnd:EventsBetweenTheTwoDates
    }
}
