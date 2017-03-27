using System;
using Aspose.Email.Mime;
using Aspose.Email.Calendar.Recurrences;
using Aspose.Email.Clients;
using Aspose.Email.Calendar;
using Aspose.Email.Clients.Smtp;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class CreateMeetingRequestWithRecurrence
    {

        public static void Run()
        {
            try
            {

                // ExStart:CreateMeetingRequestWithRecurrence
                String szUniqueId;

                // Create a mail message
                MailMessage msg1 = new MailMessage();
                msg1.To.Add("to@domain.com");
                msg1.From = new MailAddress("from@gmail.com");

                // Create appointment object
                Appointment agendaAppointment = default(Appointment);
             
                // Fill appointment object
                System.DateTime StartDate = new DateTime(2013, 12, 1, 17, 0, 0);
                System.DateTime EndDate = new DateTime(2013, 12, 31, 17, 30, 0);
                agendaAppointment = new Appointment("same place", StartDate, EndDate, msg1.From, msg1.To);

                // Create unique id as it will help to access this appointment later
                szUniqueId = Guid.NewGuid().ToString();
                agendaAppointment.UniqueId = szUniqueId;
                agendaAppointment.Description = "----------------";

                // Create a weekly reccurence pattern object
                WeeklyRecurrencePattern pattern1 = new WeeklyRecurrencePattern(14);

                // Set weekly pattern properties like days: Mon, Tue and Thu
                pattern1.StartDays = new CalendarDay[3];
                pattern1.StartDays[0] = CalendarDay.Monday;
                pattern1.StartDays[1] = CalendarDay.Tuesday;
                pattern1.StartDays[2] =CalendarDay.Thursday;
                pattern1.Interval = 1;

                // Set recurrence pattern for the appointment
                agendaAppointment.Recurrence = pattern1;

                //Attach this appointment with mail
                msg1.AlternateViews.Add(agendaAppointment.RequestApointment());

                // Create SmtpCleint
                SmtpClient client = new SmtpClient("smtp.gmail.com", 587, "your.email@gmail.com", "your.password");
                client.SecurityOptions = SecurityOptions.Auto;

                // Send mail with appointment request
                client.Send(msg1);

                // Return unique id for later usage
                // return szUniqueId;
                // ExEnd:CreateMeetingRequestWithRecurrence
            }
            catch (Exception exception)
            {
                Console.Write(exception.Message);
            }

        }
    }
}
