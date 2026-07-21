// Demonstrates how to create MAPI calendar items and add them to a PST calendar folder.

using System;
using System.IO;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.MAPI
{
    internal static class AddMapiCalendarToPst
    {
        public static void Run()
        {
            var appointment = new MapiCalendar(
                "LAKE ARGYLE WA 6743",
                "Appointment",
                "This is a very important meeting :)",
                new DateTime(2012, 10, 2, 13, 0, 0),
                new DateTime(2012, 10, 2, 14, 0, 0));

            var attendees = new MapiRecipientCollection();
            attendees.Add("ReneeAJones@armyspy.com", "Renee A. Jones", MapiRecipientType.MAPI_TO);
            attendees.Add("SzllsyLiza@dayrep.com", "Szollosy Liza", MapiRecipientType.MAPI_TO);

            var meeting = new MapiCalendar(
                "Meeting Room 3 at Office Headquarters",
                "Meeting",
                "Please confirm your availability.",
                new DateTime(2012, 10, 2, 13, 0, 0),
                new DateTime(2012, 10, 2, 14, 0, 0),
                "CharlieKhan@dayrep.com",
                attendees);

            var path = Data.Out/"AddMapiCalendarToPST_out.pst";

            if (File.Exists(path))
                File.Delete(path);

            using (var pst = PersonalStorage.Create(path, FileFormatVersion.Unicode))
            {
                // A predefined folder makes Outlook treat the items as calendar entries.
                var calendarFolder = pst.CreatePredefinedFolder("Calendar", StandardIpmFolder.Appointments);
                calendarFolder.AddMapiMessageItem(appointment);
                calendarFolder.AddMapiMessageItem(meeting);

                // An appointment has no attendees; adding them is what makes it a meeting.
                Console.WriteLine($"{appointment.Subject}: {appointment.Recipients.Count} attendee(s)");
                Console.WriteLine($"{meeting.Subject}: {meeting.Recipients.Count} attendee(s)");
                Console.WriteLine($"\nAdded {calendarFolder.ContentCount} item(s) to {path}");
            }
        }
    }
}
