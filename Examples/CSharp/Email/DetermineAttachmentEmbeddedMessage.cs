using System;
using System.IO;
using Aspose.Email.AntiSpam;
using Aspose.Email.Mail;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email
{
    internal class DetermineAttachmentEmbeddedMessage
    {
        public static void Run()
        {
            // ExStart:DetermineAttachmentEmbeddedMessage
            string dataDir = RunExamples.GetDataDir_Email() + "EmailWithAttandEmbedded.eml";
            MailMessage eml = MailMessage.Load(dataDir);

            if (eml.Attachments[0].IsEmbeddedMessage)
                Console.WriteLine("Attachment is an embedded message");
            // ExEnd:DetermineAttachmentEmbeddedMessage
        }
    }
}
