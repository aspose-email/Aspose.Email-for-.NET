using Aspose.Email.Mapi;
using Aspose.Email.Mime;
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
    class CreateDraftAppointmentFromText
    {
        public static void Run()
        {
            // ExStart:CreateDraftAppointmentFromText
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_KnowledgeBase();
            string ical = @"BEGIN:VCALENDAR
METHOD:PUBLISH
PRODID:-//Aspose Ltd//iCalender Builder (v3.0)//EN
VERSION:2.0
BEGIN:VEVENT
ATTENDEE;CN=test@gmail.com:mailto:test@gmail.com
DTSTART:20130220T171439
DTEND:20130220T174439
DTSTAMP:20130220T161439Z
END:VEVENT
END:VCALENDAR";

            string sender = "test@gmail.com";
            string recipient = "test@email.com";
            MailMessage message = new MailMessage(sender, recipient, string.Empty, string.Empty);
            AlternateView av = AlternateView.CreateAlternateViewFromString(ical, new ContentType("text/calendar"));
            message.AlternateViews.Add(av);
            MapiMessage msg = MapiMessage.FromMailMessage(message);
            msg.Save(dataDir + "DraftAppointmentFromText_out.msg");
            // ExEnd:CreateDraftAppointmentFromText
        }
    }
}
