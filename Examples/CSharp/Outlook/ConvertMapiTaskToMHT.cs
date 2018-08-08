using Aspose.Email;
using Aspose.Email.Examples.CSharp.Email;
using Aspose.Email.Mapi;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CSharp.Outlook
{
    public class ConvertMapiTaskToMHT
    {
        public static void Run()
        {
            string dataDir = RunExamples.GetDataDir_Outlook();
            //ExStart: ConvertMapiTaskToMHT
            MapiMessage msg = MapiMessage.FromFile(dataDir + "MapiTask.msg");
            MhtSaveOptions opt = SaveOptions.DefaultMhtml;
            opt.MhtFormatOptions = MhtFormatOptions.RenderTaskFields | MhtFormatOptions.WriteHeader;

            opt.FormatTemplates.Clear();
            opt.FormatTemplates.Add(MhtTemplateName.Task.Subject, "<span class='headerLineTitle'>Subject:</span><span class='headerLineText'>{0}</span><br/>");
            opt.FormatTemplates.Add(MhtTemplateName.Task.ActualWork, "<span class='headerLineTitle'>Actual Work:</span><span class='headerLineText'>{0}</span><br/>");
            opt.FormatTemplates.Add(MhtTemplateName.Task.TotalWork, "<span class='headerLineTitle'>Total Work:</span><span class='headerLineText'>{0}</span><br/>");
            opt.FormatTemplates.Add(MhtTemplateName.Task.Status, "<span class='headerLineTitle'>Status:</span><span class='headerLineText'>{0}</span><br/>");
            opt.FormatTemplates.Add(MhtTemplateName.Task.Owner, "<span class='headerLineTitle'>Owner:</span><span class='headerLineText'>{0}</span><br/>");
            opt.FormatTemplates.Add(MhtTemplateName.Task.Priority, "<span class='headerLineTitle'>Priority:</span><span class='headerLineText'>{0}</span><br/>");

            msg.Save(dataDir + "MapiTask_out.mht", opt);
            //ExEnd: ConvertMapiTaskToMHT
        }
    }
}
