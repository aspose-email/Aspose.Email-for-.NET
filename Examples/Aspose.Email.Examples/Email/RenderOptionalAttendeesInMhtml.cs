// Demonstrates how to include - or deliberately drop - the optional attendees when a
// meeting is rendered to MHTML. Clearing a field's template removes it from the output.

using System;
using System.IO;
using Aspose.Email.Calendar;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.Email
{
    internal static class RenderOptionalAttendeesInMhtml
    {
        public static void Run()
        {
            var msg = BuildMeeting();

            var options = new MhtSaveOptions
            {
                MhtFormatOptions = MhtFormatOptions.RenderCalendarEvent | MhtFormatOptions.WriteHeader
            };

            var withAttendees = Data.Out/"RenderOptionalAttendeesInMhtml_with.mhtml";
            msg.Save(withAttendees, options);
            Console.WriteLine($"With optional attendees:    {withAttendees}");
            PrintOptionalAttendees(withAttendees);

            // An empty template drops the field from the rendered header.
            options.FormatTemplates[MhtTemplateName.OptionalAttendees] = "";

            var withoutAttendees = Data.Out/"RenderOptionalAttendeesInMhtml_without.mhtml";
            msg.Save(withoutAttendees, options);
            Console.WriteLine($"Without optional attendees: {withoutAttendees}");
            PrintOptionalAttendees(withoutAttendees);
        }

        // The shipped sample meeting has no optional attendees, so build one that does.
        private static MapiMessage BuildMeeting()
        {
            var appointment = new Appointment("Meeting Room 3",
                new DateTime(2024, 5, 10, 12, 30, 0, DateTimeKind.Utc),
                new DateTime(2024, 5, 10, 13, 30, 0, DateTimeKind.Utc),
                new MailAddress("organizer@example.com", "Olivia Organizer"),
                new MailAddress("required@example.com", "Rachel Required"))
            {
                Summary = "Project sync",
                Description = "Weekly project sync."
            };

            appointment.OptionalAttendees.Add(new MailAddress("optional1@example.com", "Oscar Optional"));
            appointment.OptionalAttendees.Add(new MailAddress("optional2@example.com", "Priya Perhaps"));

            return appointment.ToMapiMessage();
        }

        private static void PrintOptionalAttendees(string path)
        {
            foreach (var line in File.ReadLines(path))
            {
                if (line.IndexOf("Optional", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    Console.WriteLine($"  {line.Trim()}");
                    return;
                }
            }

            Console.WriteLine("  (no optional attendees line in the output)");
        }
    }
}
