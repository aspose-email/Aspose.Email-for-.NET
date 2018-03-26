using Aspose.Email;
using Aspose.Email.Calendar;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CSharp.Email
{
    public class SetParticipantStatusOfAppointmentAttendees
    {
        public static void Run()
        {
            //ExStart: SetParticipantStatusOfAppointmentAttendees
            string location = "Room 5";
            DateTime startDate = new DateTime(2011, 12, 10, 10, 12, 11),
                     endDate = new DateTime(2012, 11, 13, 13, 11, 12);
            MailAddress organizer = new MailAddress("aaa@amail.com", "Organizer");
            MailAddressCollection attendees = new MailAddressCollection();
            MailAddress attendee1 = new MailAddress("bbb@bmail.com", "First attendee");
            MailAddress attendee2 = new MailAddress("ccc@cmail.com", "Second attendee");

            attendee1.ParticipationStatus = ParticipationStatus.Accepted;
            attendee2.ParticipationStatus = ParticipationStatus.Declined;
            attendees.Add(attendee1);
            attendees.Add(attendee2);

            Appointment target = new Appointment(location, startDate, endDate, organizer, attendees);
            //ExEnd: SetParticipantStatusOfAppointmentAttendees
        }
    }
}
