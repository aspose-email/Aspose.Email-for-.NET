using Aspose.Email.Mapi;
using System;
using System.Collections.Generic;
using System.IO;
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
    class ReplaceEmbeddedMSGAttachmentContents
    {
        public static void Run()
        {
            string dataDir = RunExamples.GetDataDir_Outlook();
            string fileName = dataDir + "message3.msg";

            // ExStart:ReplaceEmbeddedMSGAttachmentContents
            var message = MapiMessage.FromFile(fileName);
            var memeoryStream = new MemoryStream();
            message.Attachments[2].Save(memeoryStream);
            var getData = MapiMessage.FromStream(memeoryStream);
            message.Attachments.Replace(1, "new 1", getData);
            // ExEnd:ReplaceEmbeddedMSGAttachmentContents

            message.Save(dataDir + "ReplaceEmbeddedMSGAttachmentContents_out.msg");
        }
    }
}
