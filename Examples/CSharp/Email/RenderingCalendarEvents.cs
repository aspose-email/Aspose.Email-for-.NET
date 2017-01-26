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
namespace Aspose.Email.Examples.CSharp.Email
{
    class RenderingCalendarEvents
    {
        public static void Run()
        {
            //ExStart: RenderingCalendarEvents
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Email();
            string fileName = "Meeting with Recurring Occurrences.msg";
            MailMessage msg = MailMessage.Load(dataDir + fileName);
            MhtSaveOptions options = new MhtSaveOptions
            {
                MhtFormatOptions = MhtFormatOptions.WriteHeader | MhtFormatOptions.RenderCalendarEvent,
            };

            //Format the output details if required - optional
            MhtMessageFormatter messageformatter = new MhtMessageFormatter();
            messageformatter.StartFormat = "<span class='headerLineTitle'>Start:</span><span class='headerLineText'>{0}</span><br/>";
            messageformatter.EndFormat = "<span class='headerLineTitle'>End:</span><span class='headerLineText'>{0}</span><br/>";
            messageformatter.RecurrenceFormat = "<span class='headerLineTitle'>Recurrence:</span><span class='headerLineText'>{0}</span><br/>";
            messageformatter.RecurrencePatternFormat = "<span class='headerLineTitle'>RecurrencePattern:</span><span class='headerLineText'>{0}</span><br/>";
            messageformatter.OrganizerFormat = "<span class='headerLineTitle'>Organizer:</span><span class='headerLineText'>{0}</span><br/>";
            messageformatter.RequiredAttendeesFormat = "<span class='headerLineTitle'>RequiredAttendees:</span><span class='headerLineText'>{0}</span><br/>";
            messageformatter.Format(msg, options.MhtFormatOptions);

            msg.Save(dataDir + "Meeting with Recurring Occurrences.mhtml", Aspose.Email.Mail.SaveOptions.DefaultMhtml);
            //ExEnd: RenderingCalendarEvents
        }
    }
}
