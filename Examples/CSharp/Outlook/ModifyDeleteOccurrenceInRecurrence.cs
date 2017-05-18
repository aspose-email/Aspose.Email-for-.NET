using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;
using System;
using System.IO;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/
namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class ModifyDeleteOccurrenceInRecurrence
    {
        public static void Run()
        {
            //ExStart: ModifyDeleteOccurrenceInRecurrence
            DateTime startDate = DateTime.Now.Date.AddHours(12);

            MapiCalendarEventRecurrence recurrence = new MapiCalendarEventRecurrence();
            MapiCalendarRecurrencePattern pattern = recurrence.RecurrencePattern = new MapiCalendarDailyRecurrencePattern();
            pattern.PatternType = MapiCalendarRecurrencePatternType.Day;
            pattern.Period = 1;
            pattern.EndType = MapiCalendarRecurrenceEndType.NeverEnd;

            DateTime exceptionDate = startDate.AddDays(1);

            // adding one exception
            pattern.Exceptions.Add(new MapiCalendarExceptionInfo
            {
                Location = "London",
                Subject = "Subj",
                OriginalStartDate = exceptionDate,
                StartDateTime = exceptionDate,
                EndDateTime = exceptionDate.AddHours(5)
            });
            pattern.ModifiedInstanceDates.Add(exceptionDate);
            // every modified instance also has to have an entry in the DeletedInstanceDates field with the original instance date.
            pattern.DeletedInstanceDates.Add(exceptionDate);

            // adding one deleted instance
            pattern.DeletedInstanceDates.Add(exceptionDate.AddDays(2));


            MapiRecipientCollection recColl = new MapiRecipientCollection();
            recColl.Add("receiver@domain.com", "receiver", MapiRecipientType.MAPI_TO);
            MapiCalendar newCal = new MapiCalendar(
                "This is Location",
                "This is Summary",
                "This is recurrence test",
                startDate,
                startDate.AddHours(3),
                "organizer@domain.com",
                recColl);
            newCal.Recurrence = recurrence;

            using (MemoryStream memory = new MemoryStream())
            {
                PersonalStorage pst = PersonalStorage.Create(memory, FileFormatVersion.Unicode);
                FolderInfo calendarFolder = pst.CreatePredefinedFolder("Calendar", StandardIpmFolder.Appointments);
                calendarFolder.AddMapiMessageItem(newCal);
            }
            //ExEnd: ModifyDeleteOccurrenceInRecurrence
        }
    }
}
