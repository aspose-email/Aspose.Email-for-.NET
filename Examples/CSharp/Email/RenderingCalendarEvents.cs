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
            MhtSaveOptions options = new MhtSaveOptions();
            {
                options.MhtFormatOptions = MhtFormatOptions.WriteHeader | MhtFormatOptions.RenderCalendarEvent;

                //Format the output details if required - optional

                //Set the display for Start Property
                if (options.FormatTemplates.ContainsKey(MhtTemplateName.Start))
                    options.FormatTemplates[MhtTemplateName.Start] = @"<span class='headerLineTitle'>Start:</span><span class='headerLineText'>{0}</span><br/>"; 
                else
                    options.FormatTemplates.Add(MhtTemplateName.Start, @"<span class='headerLineTitle'>Start:</span><span class='headerLineText'>{0}</span><br/>");

                //Set the display for End Property
                if (options.FormatTemplates.ContainsKey(MhtTemplateName.End))
                    options.FormatTemplates[MhtTemplateName.End] = @"<span class='headerLineTitle'>End:</span><span class='headerLineText'>{0}</span><br/>";
                else
                    options.FormatTemplates.Add(MhtTemplateName.End, @"<span class='headerLineTitle'>End:</span><span class='headerLineText'>{0}</span><br/>");

                //Set the display for Recurrence Property
                if (options.FormatTemplates.ContainsKey(MhtTemplateName.Recurrence))
                    options.FormatTemplates[MhtTemplateName.Recurrence] = @"<span class='headerLineTitle'>Recurrence:</span><span class='headerLineText'>{0}</span><br/>";
                else
                    options.FormatTemplates.Add(MhtTemplateName.Recurrence, @"<span class='headerLineTitle'>Recurrence:</span><span class='headerLineText'>{0}</span><br/>");

                //Set the display for RecurrencePattern Property
                if (options.FormatTemplates.ContainsKey(MhtTemplateName.RecurrencePattern))
                    options.FormatTemplates[MhtTemplateName.RecurrencePattern] = @"<span class='headerLineTitle'>RecurrencePattern:</span><span class='headerLineText'>{0}</span><br/>";
                else
                    options.FormatTemplates.Add(MhtTemplateName.RecurrencePattern, @"<span class='headerLineTitle'>RecurrencePattern:</span><span class='headerLineText'>{0}</span><br/>");

                //Set the display for Organizer Property
                if (options.FormatTemplates.ContainsKey(MhtTemplateName.Organizer))
                    options.FormatTemplates[MhtTemplateName.Organizer] = @"<span class='headerLineTitle'>Organizer:</span><span class='headerLineText'>{0}</span><br/>";
                else
                    options.FormatTemplates.Add(MhtTemplateName.Organizer, @"<span class='headerLineTitle'>Organizer:</span><span class='headerLineText'>{0}</span><br/>");

                //Set the display for RequiredAttendees Property
                if (options.FormatTemplates.ContainsKey(MhtTemplateName.RequiredAttendees))
                    options.FormatTemplates[MhtTemplateName.RequiredAttendees] = @"<span class='headerLineTitle'>RequiredAttendees:</span><span class='headerLineText'>{0}</span><br/>";
                else
                    options.FormatTemplates.Add(MhtTemplateName.RequiredAttendees, @"<span class='headerLineTitle'>RequiredAttendees:</span><span class='headerLineText'>{0}</span><br/>");
            };

            msg.Save(dataDir + "Meeting with Recurring Occurrences.mhtml", options);
            //ExEnd: RenderingCalendarEvents
        }
    }
}
