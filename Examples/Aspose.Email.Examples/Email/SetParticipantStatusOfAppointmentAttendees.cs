// Demonstrates how to set participation status (Accepted/Declined) for attendees
// when creating an Appointment, and save it as an ICS file.

using Aspose.Email.Calendar;
using System;

namespace Aspose.Email.Examples.Email
{
    internal static class SetParticipantStatusOfAppointmentAttendees
    {
        public static void Run()
        {
            const string location = "Room 5";
            var startDate = new DateTime(2011, 12, 10, 10, 12, 11);
            var endDate = new DateTime(2012, 11, 13, 13, 11, 12);

            var organizer = new MailAddress("aaa@amail.com", "Organizer");
            var attendee1 = new MailAddress("bbb@bmail.com", "First attendee")
                { ParticipationStatus = ParticipationStatus.Accepted };
            var attendee2 = new MailAddress("ccc@cmail.com", "Second attendee")
                { ParticipationStatus = ParticipationStatus.Declined };

            var attendees = new MailAddressCollection { attendee1, attendee2 };

            var appointment = new Appointment(location, startDate, endDate, organizer, attendees);

            var outputPath = Data.Out/"SetParticipantStatus_out.ics";
            appointment.Save(outputPath, AppointmentSaveFormat.Ics);

            // The status is written into the ICS as the PARTSTAT parameter of each ATTENDEE.
            foreach (var attendee in appointment.Attendees)
                Console.WriteLine($"{attendee.DisplayName,-16}: {attendee.ParticipationStatus}");

            Console.WriteLine($"\nSaved to {outputPath}");
        }
    }
}
