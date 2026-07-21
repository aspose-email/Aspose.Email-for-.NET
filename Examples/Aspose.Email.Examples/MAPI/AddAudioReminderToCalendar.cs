// Demonstrates how to create a calendar item with an audio file reminder and save it as ICS.

using System;
using Aspose.Email.Mapi;
using Aspose.Email.Calendar;

namespace Aspose.Email.Examples.MAPI
{
    internal static class AddAudioReminderToCalendar
    {
        public static void Run()
        {
            var app = new Appointment("Home", DateTime.Now.AddHours(1), DateTime.Now.AddHours(1),
                "organizer@domain.com", "attendee@gmail.com");

            var msg = new MailMessage();
            msg.AddAlternateView(app.RequestApointment());

            var mapi = MapiMessage.FromMailMessage(msg);
            var calendar = (MapiCalendar)mapi.ToMapiMessageItem();

            calendar.ReminderSet = true;
            calendar.ReminderDelta = 58; // minutes before event start

            // What makes this an audio reminder: the sound file Outlook plays.
            calendar.ReminderFileParameter = Data.Mapi/"Alarm01.wav";

            var outputPath = Data.Out/"calendarWithAudioReminder_out.ics";
            calendar.Save(outputPath, AppointmentSaveFormat.Ics);

            Console.WriteLine($"Reminder {calendar.ReminderDelta} minutes before start.");
            Console.WriteLine($"Sound: {calendar.ReminderFileParameter}");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}
