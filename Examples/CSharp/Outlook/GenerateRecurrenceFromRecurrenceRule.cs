using System;
using Aspose.Email.Mapi;

/* This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET 
   API reference when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq 
   for more information. If you do not wish to use NuGet, you can manually download 
   Aspose.Email for .NET API from http://www.aspose.com/downloads, 
   install it and then add its reference to this project. For any issues, questions or suggestions 
   please feel free to contact us using http://www.aspose.com/community/forums/default.aspx            
*/

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class GenerateRecurrenceFromRecurrenceRule
    {
        public static void Run()
        {
            // ExStart:GenerateRecurrenceFromRecurrenceRule
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Outlook();

            DateTime startDate = new DateTime(2015, 7, 16);
            DateTime endDate = new DateTime(2015, 8, 1);
            MapiCalendar app1 = new MapiCalendar("test location", "test summary", "test description", startDate, endDate);

            app1.StartDate = startDate;
            app1.EndDate = endDate;

            string pattern = "DTSTART;TZID=Europe/London:20150831T080000\r\nDTEND;TZID=Europe/London:20150831T083000\r\nRRULE:FREQ=DAILY;INTERVAL=1;COUNT=7\r\nEXDATE:20150831T070000Z,20150904T070000Z";
            app1.Recurrence.RecurrencePattern = MapiCalendarRecurrencePatternFactory.FromString(pattern);
            // ExEnd:GenerateRecurrenceFromRecurrenceRule
        }         
    }
}
