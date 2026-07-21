// Demonstrates how to convert a MAPI task MSG to MHTML using custom field templates.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class ConvertMapiTaskToMht
    {
        public static void Run()
        {
            var msg = MapiMessage.Load(Data.Mapi/"MapiTask.msg");
            var opt = SaveOptions.DefaultMhtml;

            // RenderTaskFields is what puts the task details into the output; the
            // templates below then control the markup of each field.
            opt.MhtFormatOptions = MhtFormatOptions.RenderTaskFields | MhtFormatOptions.WriteHeader;

            // Clearing first means only the fields listed below are rendered.
            opt.FormatTemplates.Clear();
            opt.FormatTemplates.Add(MhtTemplateName.Task.Subject,    "<span class='headerLineTitle'>Subject:</span><span class='headerLineText'>{0}</span><br/>");
            opt.FormatTemplates.Add(MhtTemplateName.Task.ActualWork, "<span class='headerLineTitle'>Actual Work:</span><span class='headerLineText'>{0}</span><br/>");
            opt.FormatTemplates.Add(MhtTemplateName.Task.TotalWork,  "<span class='headerLineTitle'>Total Work:</span><span class='headerLineText'>{0}</span><br/>");
            opt.FormatTemplates.Add(MhtTemplateName.Task.Status,     "<span class='headerLineTitle'>Status:</span><span class='headerLineText'>{0}</span><br/>");
            opt.FormatTemplates.Add(MhtTemplateName.Task.Owner,      "<span class='headerLineTitle'>Owner:</span><span class='headerLineText'>{0}</span><br/>");
            opt.FormatTemplates.Add(MhtTemplateName.Task.Priority,   "<span class='headerLineTitle'>Priority:</span><span class='headerLineText'>{0}</span><br/>");

            var outputPath = Data.Out/"MapiTask_out.mht";
            msg.Save(outputPath, opt);

            Console.WriteLine($"Task: {msg.Subject}");
            Console.WriteLine($"Rendered {opt.FormatTemplates.Count} task field(s).");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}
