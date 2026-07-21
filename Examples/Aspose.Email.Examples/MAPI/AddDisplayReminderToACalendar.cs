// Demonstrates how to create a calendar item with a display reminder and save it as ICS.

using System;
using Aspose.Email.Mapi;
using Aspose.Email.Calendar;

namespace Aspose.Email.Examples.MAPI
{
    internal static class AddDisplayReminderToACalendar
    {
        public static void Run()
        {
            var app = new Appointment("Home", DateTime.Now.AddHours(1), DateTime.Now.AddHours(1),
                "organizer@domain.com", "attendee@gmail.com");

            var msg = new MailMessage();
            msg.AddAlternateView(app.RequestApointment());

            var mapi = MapiMessage.FromMailMessage(msg);
            var calendar = (MapiCalendar)mapi.ToMapiMessageItem();

            // A display reminder needs no sound file - ReminderSet plus the delta is
            // enough for Outlook to pop the reminder up.
            calendar.ReminderSet = true;
            calendar.ReminderDelta = 45; // minutes before event start

            var outputPath = Data.Out/"calendarWithDisplayReminder.ics";
            calendar.Save(outputPath, AppointmentSaveFormat.Ics);

            Console.WriteLine($"Reminder {calendar.ReminderDelta} minutes before start.");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}
