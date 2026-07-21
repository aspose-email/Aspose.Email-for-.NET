// Demonstrates how to create MAPI calendar items and add them to a subfolder of the
// PST calendar folder, rather than to the calendar folder itself.

using System;
using System.IO;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.MAPI
{
    internal static class CreateNewMapiCalendarAndAddToCalendarSubfolder
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

            var meeting = new MapiCalendar(
                "Meeting Room 3 at Office Headquarters",
                "Meeting",
                "Please confirm your availability.",
                new DateTime(2012, 10, 3, 13, 0, 0),
                new DateTime(2012, 10, 3, 14, 0, 0),
                "CharlieKhan@dayrep.com",
                attendees);

            var path = Data.Out/"CreateNewMapiCalendarAndAddToCalendarSubfolder_out.pst";

            if (File.Exists(path))
                File.Delete(path);

            using (var pst = PersonalStorage.Create(path, FileFormatVersion.Unicode))
            {
                // The predefined folder is what makes Outlook treat its contents as
                // calendar items; a subfolder under it inherits that behaviour, which is
                // how a second calendar is added alongside the default one.
                var calendarFolder = pst.CreatePredefinedFolder("Calendar", StandardIpmFolder.Appointments);
                var subFolder = calendarFolder.AddSubFolder("Team calendar");

                foreach (var item in new[] { appointment, meeting })
                {
                    subFolder.AddMapiMessageItem(item);
                    Console.WriteLine($"{item.Subject,-12} {item.StartDate:g} at {item.Location}");
                }

                Console.WriteLine($"\nAdded {subFolder.ContentCount} item(s) to " +
                                  $"\"{calendarFolder.DisplayName}\\{subFolder.DisplayName}\"");
                Console.WriteLine($"Saved to {path}");
            }
        }
    }
}
