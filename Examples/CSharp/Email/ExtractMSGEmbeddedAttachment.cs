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
    class ExtractMSGEmbeddedAttachment
    {
        public static void Run()
        {
            string dataDir = RunExamples.GetDataDir_Email();
            ExtractInlineAttachments(dataDir);
        }

        // ExStart:ExtractMSGEmbeddedAttachment
        public static void ExtractInlineAttachments(string dataDir)
        {
            MapiMessage message = MapiMessage.FromFile(dataDir + "MSG file with RTF Formatting.msg");
            MapiAttachmentCollection attachments = message.Attachments;
            foreach (MapiAttachment attachment in attachments)
            {
        
                if(IsAttachmentInline(attachment))
                {
                    try
                    {
                        SaveAttachment(attachment, new Guid().ToString());
                    }
                    catch (Exception ex)
                    {                        
                       Console.WriteLine(ex.Message);
                    }
                }
            }
        }

        static bool IsAttachmentInline(MapiAttachment attachment)
        {
            foreach (MapiProperty property in attachment.ObjectData.Properties.Values)
            {
                if (property.Name == "\x0003ObjInfo")
                {
                    ushort odtPersist1 = BitConverter.ToUInt16(property.Data, 0);
                    return (odtPersist1 & (1 << (7 - 1))) == 0;
                }
            }
            return false;
        }

        static void SaveAttachment(MapiAttachment attachment, string fileName)
        {
            foreach (MapiProperty property in attachment.ObjectData.Properties.Values)
            {
                if (property.Name == "Package")
                {
                    using (FileStream fs = new FileStream(fileName, FileMode.Create, FileAccess.Write))
                    {
                        fs.Write(property.Data, 0, property.Data.Length);
                    }
                }
            }
        }
        // ExEnd:ExtractMSGEmbeddedAttachment
    }
}
