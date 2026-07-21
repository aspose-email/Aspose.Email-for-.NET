// Demonstrates how to change one occurrence of a recurring appointment and delete
// another, without touching the rest of the series.

using System;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.MAPI
{
    internal static class ModifyDeleteOccurrenceInRecurrence
    {
        public static void Run()
        {
            var startDate = DateTime.Now.Date.AddHours(12);

            var recurrence = new MapiCalendarEventRecurrence();
            var pattern = recurrence.RecurrencePattern = new MapiCalendarDailyRecurrencePattern
            {
                PatternType = MapiCalendarRecurrencePatternType.Day,
                Period = 1,
                EndType = MapiCalendarRecurrenceEndType.NeverEnd
            };

            var exceptionDate = startDate.AddDays(1);

            // An exception describes what one occurrence looks like when it differs from
            // the series - here it runs longer and in a different place.
            pattern.Exceptions.Add(new MapiCalendarExceptionInfo
            {
                Location = "London",
                Subject = "Subj",
                OriginalStartDate = exceptionDate,
                StartDateTime = exceptionDate,
                EndDateTime = exceptionDate.AddHours(5)
            });
            pattern.ModifiedInstanceDates.Add(exceptionDate);

            // Every modified instance also needs an entry in DeletedInstanceDates with the
            // original instance date: the original is removed and the exception replaces it.
            pattern.DeletedInstanceDates.Add(exceptionDate);

            // A date listed only here, with no exception, is simply dropped from the series.
            pattern.DeletedInstanceDates.Add(exceptionDate.AddDays(2));

            var recipients = new MapiRecipientCollection();
            recipients.Add("receiver@domain.com", "receiver", MapiRecipientType.MAPI_TO);

            var newCal = new MapiCalendar(
                "This is Location",
                "This is Summary",
                "This is recurrence test",
                startDate,
                startDate.AddHours(3),
                "organizer@domain.com",
                recipients)
            {
                Recurrence = recurrence
            };

            var pstPath = Data.Out/"ModifyDeleteOccurrenceInRecurrence_out.pst";

            using (var pst = PersonalStorage.Create(pstPath, FileFormatVersion.Unicode))
            {
                var calendarFolder = pst.CreatePredefinedFolder("Calendar", StandardIpmFolder.Appointments);
                calendarFolder.AddMapiMessageItem(newCal);
            }

            Console.WriteLine($"Series starts:      {startDate:g}, repeats every day");
            Console.WriteLine($"Modified occurrence: {exceptionDate:g} moved to London, 5 hours long");
            Console.WriteLine($"Deleted occurrence:  {exceptionDate.AddDays(2):g}");
            Console.WriteLine($"Saved to {pstPath}");
        }
    }
}
