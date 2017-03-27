using System;
using Aspose.Email.Mime;
using Aspose.Email.Calendar;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class SetReminderByAddingTags
    {
        public static void Run()
        {
            // ExStart:SetReminderByAddingTags
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Outlook();

            string location = "Meeting Location: Room 5";
            DateTime startDate = new DateTime(1997, 3, 18, 18, 30, 00),
            endDate = new DateTime(1997, 3, 18, 19, 30, 00);
            MailAddress organizer = new MailAddress("aaa@amail.com", "Organizer");
            MailAddressCollection attendees = new MailAddressCollection();
            attendees.Add(new MailAddress("bbb@bmail.com", "First attendee"));

            Appointment target = new Appointment(location, startDate, endDate, organizer, attendees);

            // Audio alarm that will sound at a precise time and
            // Repeat 4 more times at 15 minute intervals:
            AppointmentReminder audioReminder = new AppointmentReminder();
            audioReminder.Trigger = new ReminderTrigger(new DateTime(1997, 3, 17, 13, 30, 0, DateTimeKind.Utc));
            audioReminder.Repeat = 4;
            audioReminder.Duration = new ReminderDuration(new TimeSpan(0, 15, 0));
            audioReminder.Action = ReminderAction.Audio;
            ReminderAttachment attach = new ReminderAttachment(new Uri("ftp://Host.com/pub/sounds/bell-01.aud"));
            audioReminder.Attachments.Add(attach);
            target.Reminders.Add(audioReminder);


            // Display alarm that will trigger 30 minutes before the
            // Scheduled start of the event it is
            // Associated with and will repeat 2 more times at 15 minute intervals:
            AppointmentReminder displayReminder = new AppointmentReminder();
            ReminderDuration dur = new ReminderDuration(new TimeSpan(0, -30, 0));
            displayReminder.Trigger = new ReminderTrigger(dur, ReminderRelated.Start);
            displayReminder.Repeat = 2;
            displayReminder.Duration = new ReminderDuration(new TimeSpan(0, 15, 0));
            displayReminder.Action = ReminderAction.Display;
            displayReminder.Description = "Breakfast meeting with executive team at 8:30 AM EST";
            target.Reminders.Add(displayReminder);

            // Email alarm that will trigger 2 days before the
            // Scheduled due date/time. It does not
            // Repeat. The email has a subject, body and attachment link.
            AppointmentReminder emailReminder = new AppointmentReminder();
            ReminderDuration dur1 = new ReminderDuration(new TimeSpan(-2, 0, 0, 0));
            emailReminder.Trigger = new ReminderTrigger(dur1, ReminderRelated.Start);
            ReminderAttendee attendee = new ReminderAttendee("john_doe@host.com");
            emailReminder.Attendees.Add(attendee);
            emailReminder.Action = ReminderAction.Email;
            emailReminder.Summary = "REMINDER: SEND AGENDA FOR WEEKLY STAFF MEETING";
            emailReminder.Description = @"A draft agenda needs to be sent out to the attendees to the weekly managers meeting (MGR-LIST). Attached is a pointer the document template for the agenda file.";
            ReminderAttachment attach1 = new ReminderAttachment(new Uri("http://Host.com/templates/agenda.doc"));
            emailReminder.Attachments.Add(attach1);
            target.Reminders.Add(emailReminder);

            // Procedural alarm that will trigger at a precise date/time
            // And will repeat 23 more times at one hour intervals. The alarm will
            // Invoke a procedure file.
            AppointmentReminder procReminder = new AppointmentReminder();
            procReminder.Trigger = new ReminderTrigger(new DateTime(1998, 1, 1, 5, 0, 0, DateTimeKind.Utc));
            procReminder.Repeat = 23;
            procReminder.Duration = new ReminderDuration(new TimeSpan(1, 0, 0));
            procReminder.Action = ReminderAction.Procedure;
            ReminderAttachment attach2 = new ReminderAttachment(new Uri("ftp://Host.com/novo-procs/felizano.exe"));
            procReminder.Attachments.Add(attach2);
            target.Reminders.Add(procReminder);
            target.Save(dataDir + "savedFile_out.ics");
            //ExEnd:SetReminderByAddingTags
        }
    }
}