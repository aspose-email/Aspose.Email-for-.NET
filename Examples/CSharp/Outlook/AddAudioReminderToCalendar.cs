using System;
using Aspose.Email.Mapi;
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
    class AddAudioReminderToCalendar
    {
        public static void Run()
        {
            // ExStart:AddAudioReminderToCalendar
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Outlook();
            Appointment app = new Appointment("Home", DateTime.Now.AddHours(1), DateTime.Now.AddHours(1), "organizer@domain.com", "attendee@gmail.com");

            MailMessage msg = new MailMessage();
            msg.AddAlternateView(app.RequestApointment());
            MapiMessage mapi = MapiMessage.FromMailMessage(msg);
            MapiCalendar calendar = (MapiCalendar)mapi.ToMapiMessageItem();

            // Set calendar properties
            calendar.ReminderSet = true;
            calendar.ReminderDelta = 58;//58 min before start of event
            calendar.ReminderFileParameter = dataDir + "Alarm01.wav";
            string savedFile = (dataDir + "calendarWithAudioReminder_out.ics");
            calendar.Save(savedFile, AppointmentSaveFormat.Ics);
            // ExEnd:AddAudioReminderToCalendar
        }
    }
}