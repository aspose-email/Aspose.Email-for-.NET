// Demonstrates how to save an appointment as an iCalendar (.ics) file and read it back.
//
// An .ics file is what calendar programs exchange: attach it to a message, offer it for
// download, or import it into Outlook, Google Calendar and others. MeetingRequests shows
// how to send the same appointment as a meeting request instead.

using System;
using Aspose.Email.Calendar;

namespace Aspose.Email.Examples.SMTP
{
    internal static class AppointmentInICSFormat
    {
        public static void Run()
        {
            var outputPath = Data.Out/"AppointmentInICSFormat_out.ics";

            var appointment = new Appointment(
                "Meeting Room 3 at Office Headquarters",  // location
                "Monthly Meeting",                        // summary
                "Please confirm your availability.",      // description
                new DateTime(2026, 2, 8, 13, 0, 0),       // start
                new DateTime(2026, 2, 8, 14, 0, 0),       // end
                "organizer@example.com",                  // organizer
                "attendee@example.com");                  // attendees

            appointment.CreatedDate = new DateTime(2026, 1, 15, 0, 0, 0, DateTimeKind.Utc);
            appointment.LastModifiedDate = new DateTime(2026, 1, 16, 0, 0, 0, DateTimeKind.Utc);

            appointment.Save(outputPath, AppointmentSaveFormat.Ics);
            Console.WriteLine($"Saved to {outputPath}");

            var loaded = Appointment.Load(outputPath);

            Console.WriteLine("\nRead back:");
            Console.WriteLine($"  summary:       {loaded.Summary}");
            Console.WriteLine($"  location:      {loaded.Location}");
            Console.WriteLine($"  description:   {loaded.Description}");
            Console.WriteLine($"  start:         {loaded.StartDate}");
            Console.WriteLine($"  end:           {loaded.EndDate}");
            Console.WriteLine($"  organizer:     {loaded.Organizer}");
            Console.WriteLine($"  attendees:     {loaded.Attendees}");
            Console.WriteLine($"  created:       {loaded.CreatedDate}");
            Console.WriteLine($"  last modified: {loaded.LastModifiedDate}");
        }
    }
}
