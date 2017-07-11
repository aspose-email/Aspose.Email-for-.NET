using System;
using Aspose.Email.Tools;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email
{
    public class PreserveEmbeddedMSGFormatDuringLoad
    {
        public static void Run()
        {
            //ExStart: PreserveEmbeddedMSGFormatDuringLoad
            string dataDir = RunExamples.GetDataDir_Email();

            MailMessage mail = MailMessage.Load(dataDir + "tnefWithMsgInside.eml", new EmlLoadOptions() { PreserveEmbeddedMessageFormat = true });

            FileFormatType fileFormat = FileFormatUtil.DetectFileFormat(mail.Attachments[0].ContentStream).FileFormatType;

            Console.WriteLine("Embedded message file format: " + fileFormat);
            //ExEnd: PreserveEmbeddedMSGFormatDuringLoad
        }
    }
}
