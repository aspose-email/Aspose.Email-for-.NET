// Demonstrates the meeting-specific properties of a MapiCalendar: who is invited, how
// the time shows in the organiser's calendar, and whether replies are expected.
//
// One trap worth knowing: MapiCalendar.Save(path) with no format writes a file that
// cannot be read back. Pass AppointmentSaveFormat.Msg (or MapiCalendarMsgSaveOptions)
// so the calendar-specific properties are written properly.

using System;
using Aspose.Email.Calendar;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class SetCalendarAttendeesAndBusyStatus
    {
        public static void Run()
        {
            ReadExisting();
            Console.WriteLine();
            BuildNew();
        }

        private static void ReadExisting()
        {
            var calendar = (MapiCalendar)MapiMessage.Load(Data.Mapi/"Test Meeting.msg").ToMapiMessageItem();

            Console.WriteLine($"Meeting:     {calendar.Subject}");
            Console.WriteLine($"Busy status: {calendar.BusyStatus}");
            Console.WriteLine($"Location:    {calendar.Location}");

            var attendees = calendar.Attendees;
            Console.WriteLine($"Response requested: {attendees.ResponseRequested}");
            Console.WriteLine($"New times may be proposed: {!attendees.NotAllowPropose}");

            Console.WriteLine($"Invited ({attendees.AppointmentRecipients.Count}):");
            foreach (var recipient in attendees.AppointmentRecipients)
            {
                Console.WriteLine($"  {recipient.DisplayName} <{recipient.EmailAddress}> " +
                                  $"{recipient.RecipientType}, replied: {recipient.RecipientTrackStatus}");
            }

            // Recipients that cannot be sent to - rooms and resources typically land here.
            Console.WriteLine($"Unsendable recipients: {attendees.AppointmentUnsendableRecipients.Count}");
        }

        private static void BuildNew()
        {
            var calendar = new MapiCalendar(
                "Meeting Room 3",
                "Project sync",
                "Weekly project sync.",
                new DateTime(2024, 5, 10, 12, 30, 0, DateTimeKind.Utc),
                new DateTime(2024, 5, 10, 13, 30, 0, DateTimeKind.Utc))
            {
                // How the slot appears to whoever looks at the organiser's calendar.
                BusyStatus = MapiCalendarBusyStatus.OutOfOffice,
                ClientIntent = MapiCalendarClientIntent.Manager,
                Organizer = new MapiElectronicAddress
                {
                    EmailAddress = "organizer@example.com",
                    DisplayName = "Olivia Organizer"
                }
            };

            calendar.Attendees = new MapiCalendarAttendees
            {
                ResponseRequested = true,
                NotAllowPropose = true
            };
            calendar.Attendees.AppointmentRecipients.Add(
                "required@example.com", "SMTP", "Rachel Required", MapiRecipientType.MAPI_TO);
            calendar.Attendees.AppointmentRecipients.Add(
                "optional@example.com", "SMTP", "Oscar Optional", MapiRecipientType.MAPI_CC);

            var outputPath = Data.Out/"SetCalendarAttendeesAndBusyStatus_out.msg";
            calendar.Save(outputPath, AppointmentSaveFormat.Msg);

            // Read it back to show the properties survived the save.
            var reloaded = (MapiCalendar)MapiMessage.Load(outputPath).ToMapiMessageItem();

            Console.WriteLine($"Created:     {reloaded.Subject}");
            Console.WriteLine($"Busy status: {reloaded.BusyStatus}");
            Console.WriteLine($"Intent:      {reloaded.ClientIntent}");
            Console.WriteLine($"Attendees:   {reloaded.Attendees.AppointmentRecipients.Count}");
            Console.WriteLine($"Response requested: {reloaded.Attendees.ResponseRequested}");
            Console.WriteLine($"\nSaved to {outputPath}");
        }
    }
}
