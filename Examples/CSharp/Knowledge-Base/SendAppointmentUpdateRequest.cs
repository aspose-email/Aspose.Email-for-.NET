using Aspose.Email.Mail;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Knowledge.Base
{
    class SendAppointmentUpdateRequest
    {
        public static void Run()
        {
            try
            {
                EMAIL_UPDATE_RECURRENCE("szUniqueId");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        // ExStart:SendAppointmentUpdateRequest
        static public void EMAIL_UPDATE_RECURRENCE(String szUniqueId)
        {
            System.DateTime StartDate = new DateTime(2013, 12, 12, 17, 0, 0);
            System.DateTime EndDate = new DateTime(2013, 12, 12, 17, 30, 0);
            Appointment appUpdate = new Appointment("Different Place", StartDate,
            EndDate, "newcustomeronnet@gmail.com", "kashif.iqbal@aspose.com");
            appUpdate.UniqueId = szUniqueId;
            Aspose.Email.Mail.Calendaring.WeeklyRecurrencePattern pattern3 = (Aspose.Email.Mail.Calendaring.WeeklyRecurrencePattern)appUpdate.RecurrencePattern;
            appUpdate.Summary = "update meeting request summary";
            appUpdate.Description = "update";
            MailMessage msgUpdate = new MailMessage("newcustomeronnet@gmail.com", "kashif.iqbal@aspose.com", "06 - test email - update meeting request", "test email");
            msgUpdate.AddAlternateView(appUpdate.UpdateAppointment());
            SmtpClient smtp = new SmtpClient("server.domain.com", 587, "username", "password");
            smtp.Send(msgUpdate);
        }
        // ExEnd:SendAppointmentUpdateRequest
    }
}
