// Demonstrates how to render a calendar event email to MHTML format with
// custom HTML templates for event properties such as Start, End, Recurrence, and Organizer.

using System;

namespace Aspose.Email.Examples.Email
{
    internal static class RenderingCalendarEvents
    {
        public static void Run()
        {
            var eml = MailMessage.Load(Data.Email/"Meeting with Recurring Occurrences.msg");

            // RenderCalendarEvent is what makes the event details appear in the output at
            // all; the templates below then control how each field is marked up.
            var options = new MhtSaveOptions
            {
                MhtFormatOptions = MhtFormatOptions.WriteHeader | MhtFormatOptions.RenderCalendarEvent
            };

            SetTemplate(options, MhtTemplateName.Start,
                "<span class='headerLineTitle'>Start:</span><span class='headerLineText'>{0}</span><br/>");
            SetTemplate(options, MhtTemplateName.End,
                "<span class='headerLineTitle'>End:</span><span class='headerLineText'>{0}</span><br/>");
            SetTemplate(options, MhtTemplateName.Recurrence,
                "<span class='headerLineTitle'>Recurrence:</span><span class='headerLineText'>{0}</span><br/>");
            SetTemplate(options, MhtTemplateName.RecurrencePattern,
                "<span class='headerLineTitle'>RecurrencePattern:</span><span class='headerLineText'>{0}</span><br/>");
            SetTemplate(options, MhtTemplateName.Organizer,
                "<span class='headerLineTitle'>Organizer:</span><span class='headerLineText'>{0}</span><br/>");
            SetTemplate(options, MhtTemplateName.RequiredAttendees,
                "<span class='headerLineTitle'>RequiredAttendees:</span><span class='headerLineText'>{0}</span><br/>");

            var outputPath = Data.Out/"Meeting with Recurring Occurrences.mhtml";
            eml.Save(outputPath, options);

            Console.WriteLine($"Rendered: {eml.Subject}");
            Console.WriteLine($"Customised {options.FormatTemplates.Count} field template(s).");
            Console.WriteLine($"Saved to {outputPath}");
        }

        private static void SetTemplate(MhtSaveOptions options, string key, string template)
        {
            if (options.FormatTemplates.ContainsKey(key))
                options.FormatTemplates[key] = template;
            else
                options.FormatTemplates.Add(key, template);
        }
    }
}
