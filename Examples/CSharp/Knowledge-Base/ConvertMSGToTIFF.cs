using Aspose.Words;
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

namespace Aspose.Email.Examples.CSharp.Email.Knowledge.Base
{
    class ConvertMSGToTIFF
    {
        public static void Run()
        {
            // ExStart:ConvertMSGToTIFF
            // The path to the File directory  and load the MSG file using Aspose.Email for .NET
            string dataDir = RunExamples.GetDataDir_KnowledgeBase();
            MailMessage msg = MailMessage.Load(dataDir + "message3.msg", new MsgLoadOptions());

            // Convert MSG to MHTML and save to stream
            MemoryStream msgStream = new MemoryStream();
            msg.Save(msgStream, SaveOptions.DefaultMhtml);
            msgStream.Position = 0;

            // Load the MHTML stream using Aspose.Words for .NET and Save the document as TIFF image
            Document msgDocument = new Document(msgStream);
            msgDocument.Save(dataDir + "Outlook-Aspose_out.tif", SaveFormat.Tiff);
            // ExEnd:ConvertMSGToTIFF
        }
    }
}
