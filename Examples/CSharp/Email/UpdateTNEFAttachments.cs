using Aspose.Email.Mime;
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
    class UpdateTNEFAttachments
    {
        public static void Run()
        {
            string dataDir = RunExamples.GetDataDir_Email();
            TestUpdateResources(dataDir);
        }

        // ExStart:UpdateTNEFAttachments
        public static void TestUpdateResources(string dataDir)
        {
            string fileName = dataDir + "tnefEML1.eml";
            string imgFileName = dataDir + "Untitled.jpg";
            string outFileName = dataDir + "UpdateTNEFAttachments_out.eml";
            MailMessage originalMailMessage = MailMessage.Load(fileName);
            UpdateResources(originalMailMessage, imgFileName);
            EmlSaveOptions emlSo = new EmlSaveOptions(MailMessageSaveType.EmlFormat);
            emlSo.FileCompatibilityMode = FileCompatibilityMode.PreserveTnefAttachments;
            originalMailMessage.Save(outFileName, emlSo);
        }

        private static void UpdateResources(MailMessage msg, string imgFileName)
        {
            for (int i = 0; i < msg.Attachments.Count; i++)
            {
                if ((msg.Attachments[i].ContentType.MediaType == "image/png") || (msg.Attachments[i].ContentType.MediaType == "application/octet-stream" && Path.GetExtension(msg.Attachments[i].ContentType.Name) == ".jpg"))
                {
                    msg.Attachments[i].ContentStream = new MemoryStream(File.ReadAllBytes(imgFileName));
                }
                else if ((msg.Attachments[i].ContentType.MediaType == "message/rfc822") || (msg.Attachments[i].ContentType.MediaType == "application/octet-stream" && Path.GetExtension(msg.Attachments[i].ContentType.Name) == ".msg"))
                {
                    MemoryStream ms = new MemoryStream();
                    msg.Attachments[i].Save(ms);
                    ms.Position = 0;
                    MailMessage embeddedMessage = MailMessage.Load(ms);
                    UpdateResources(embeddedMessage, imgFileName);
                    MemoryStream msProcessedEmbedded = new MemoryStream();
                    embeddedMessage.Save(msProcessedEmbedded, SaveOptions.DefaultMsgUnicode);
                    msProcessedEmbedded.Position = 0;
                    msg.Attachments[i].ContentStream = msProcessedEmbedded;
                }
            }

            foreach (LinkedResource att in msg.LinkedResources)
            {
                if (att.ContentType.MediaType == "image/png")
                    att.ContentStream = new MemoryStream(File.ReadAllBytes(imgFileName));
            }
        }
        // ExEnd:UpdateTNEFAttachments
    }
}
