using System;
using System.IO;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from https://www.nuget.org/packages/Aspose.Email/, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using https://forum.aspose.com/c/email
*/

namespace Aspose.Email.Examples.CSharp.Email
{
    class TrackDocumentConversionProgress
    {
        public static void Run()
        {
            // ExStart:1
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Email();

            var fileName = dataDir + "test.eml";
            MailMessage msg = MailMessage.Load(fileName);

            MemoryStream ms = new MemoryStream();
            EmlSaveOptions opt = new EmlSaveOptions(MailMessageSaveType.EmlFormat);
            opt.CustomProgressHandler = new ConversionProgressEventHandler(ShowEmlConversionProgress);
            msg.Save(ms, opt);
            // ExEnd:1
        }

        // ExStart:2
        private static void ShowEmlConversionProgress(ProgressEventHandlerInfo info)
        {
            int total;
            int saved;
            switch (info.EventType)
            {
                case ProgressEventType.MimeStructureCreated:
                    total = info.TotalMimePartCount;
                    saved = info.SavedMimePartCount;
                    Console.WriteLine("MimeStructureCreated - TotalMimePartCount: " + total);
                    Console.WriteLine("MimeStructureCreated - SavedMimePartCount: " + saved);
                    break;
                case ProgressEventType.MimePartSaved:
                    total = info.TotalMimePartCount;
                    saved = info.SavedMimePartCount;
                    Console.WriteLine("MimePartSaved - TotalMimePartCount: " + total);
                    Console.WriteLine("MimePartSaved - SavedMimePartCount: " + saved);
                    break;
                case ProgressEventType.SavedToStream:
                    total = info.TotalMimePartCount;
                    saved = info.SavedMimePartCount;
                    Console.WriteLine("SavedToStream - TotalMimePartCount: " + total);
                    Console.WriteLine("SavedToStream - SavedMimePartCount: " + saved);
                    break;
            }
        }
        // ExEnd:2
    }


}
