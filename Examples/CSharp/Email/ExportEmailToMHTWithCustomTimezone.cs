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

namespace Aspose.Email.Examples.CSharp.Email
{
    class ExportEmailToMHTWithCustomTimezone
    {
        public static void Run()
        {
            string dataDir = RunExamples.GetDataDir_Email();

            // ExStart:ExportEmailToMHTWithCustomTimezone
            MailMessage eml = MailMessage.Load(dataDir + "Message.eml");

            // Set the local time for message date.
            eml.TimeZoneOffset = TimeZone.CurrentTimeZone.GetUtcOffset(DateTime.Now);

            // or set custom time zone offset for message date (-0800)
            // eml.TimeZoneOffset = new TimeSpan(-8,0,0);

            // The dates will be rendered by local system time zone.
            MhtSaveOptions so = new MhtSaveOptions();
            so.MhtFormatOptions = MhtFormatOptions.WriteHeader;
            eml.Save(dataDir + "ExportEmailToMHTWithCustomTimezone_out.mhtml", so);
            // ExEnd:ExportEmailToMHTWithCustomTimezone
        }
    }
}
