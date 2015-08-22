using Aspose.Email.Mail;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Email
{
    class Program
    {
        static void Main(string[] args)
        {
            string location = "Meeting Location: Room 5";
            DateTime startDate = new DateTime(1997, 3, 18, 18, 30, 00),
                endDate = new DateTime(1997, 3, 18, 19, 30, 00);
            MailAddress organizer = new MailAddress("aaa@amail.com", "Organizer");
            MailAddressCollection attendees = new MailAddressCollection();
            attendees.Add(new MailAddress("bbb@bmail.com", "First attendee"));

            Appointment target = new Appointment(location, startDate, endDate, organizer, attendees);
            target.Save("savedFile.ics");
        }
    }
}
