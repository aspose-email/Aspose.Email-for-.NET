// Demonstrates the four kinds of reminder an appointment can carry - audio,
// display, email and procedure - and how each one is triggered.

using System;
using Aspose.Email.Calendar;

namespace Aspose.Email.Examples.MAPI
{
    internal static class SetReminderByAddingTags
    {
        public static void Run()
        {
            var startDate = new DateTime(1997, 3, 18, 18, 30, 00);
            var endDate = new DateTime(1997, 3, 18, 19, 30, 00);

            var organizer = new MailAddress("aaa@amail.com", "Organizer");
            var attendees = new MailAddressCollection { new MailAddress("bbb@bmail.com", "First attendee") };

            var target = new Appointment("Meeting Location: Room 5", startDate, endDate, organizer, attendees);

            target.Reminders.Add(CreateAudioReminder());
            target.Reminders.Add(CreateDisplayReminder());
            target.Reminders.Add(CreateEmailReminder());
            target.Reminders.Add(CreateProcedureReminder());

            var outputPath = Data.Out/"savedFile_out.ics";
            target.Save(outputPath);

            foreach (var reminder in target.Reminders)
                Console.WriteLine($"{reminder.Action,-10} reminder, repeats {reminder.Repeat} more time(s)");

            Console.WriteLine($"\nSaved {target.Reminders.Count} reminder(s) to {outputPath}");
        }

        // Triggers at an absolute time, then repeats 4 more times every 15 minutes.
        private static AppointmentReminder CreateAudioReminder()
        {
            var reminder = new AppointmentReminder
            {
                Trigger = new ReminderTrigger(new DateTime(1997, 3, 17, 13, 30, 0, DateTimeKind.Utc)),
                Repeat = 4,
                Duration = new ReminderDuration(new TimeSpan(0, 15, 0)),
                Action = ReminderAction.Audio
            };

            reminder.Attachments.Add(new ReminderAttachment(new Uri("ftp://Host.com/pub/sounds/bell-01.aud")));
            return reminder;
        }

        // A negative duration relative to Start means 30 minutes before the event.
        private static AppointmentReminder CreateDisplayReminder()
        {
            return new AppointmentReminder
            {
                Trigger = new ReminderTrigger(new ReminderDuration(new TimeSpan(0, -30, 0)), ReminderRelated.Start),
                Repeat = 2,
                Duration = new ReminderDuration(new TimeSpan(0, 15, 0)),
                Action = ReminderAction.Display,
                Description = "Breakfast meeting with executive team at 8:30 AM EST"
            };
        }

        // Fires once, two days before the event, and mails the attendee.
        private static AppointmentReminder CreateEmailReminder()
        {
            var reminder = new AppointmentReminder
            {
                Trigger = new ReminderTrigger(new ReminderDuration(new TimeSpan(-2, 0, 0, 0)), ReminderRelated.Start),
                Action = ReminderAction.Email,
                Summary = "REMINDER: SEND AGENDA FOR WEEKLY STAFF MEETING",
                Description = "A draft agenda needs to be sent out to the attendees to the weekly " +
                              "managers meeting (MGR-LIST). Attached is a pointer the document " +
                              "template for the agenda file."
            };

            reminder.Attendees.Add(new ReminderAttendee("john_doe@host.com"));
            reminder.Attachments.Add(new ReminderAttachment(new Uri("http://Host.com/templates/agenda.doc")));
            return reminder;
        }

        // Invokes a program at an absolute time, repeating hourly.
        private static AppointmentReminder CreateProcedureReminder()
        {
            var reminder = new AppointmentReminder
            {
                Trigger = new ReminderTrigger(new DateTime(1998, 1, 1, 5, 0, 0, DateTimeKind.Utc)),
                Repeat = 23,
                Duration = new ReminderDuration(new TimeSpan(1, 0, 0)),
                Action = ReminderAction.Procedure
            };

            reminder.Attachments.Add(new ReminderAttachment(new Uri("ftp://Host.com/novo-procs/felizano.exe")));
            return reminder;
        }
    }
}
