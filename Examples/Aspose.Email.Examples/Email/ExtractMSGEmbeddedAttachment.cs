// Demonstrates how to identify inline OLE attachments in an MSG file and save them to disk.

using System;
using System.IO;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.Email
{
    internal static class ExtractMsgEmbeddedAttachment
    {
        public static void Run()
        {
            var eml = MapiMessage.Load(Data.Email/"MSG file with RTF Formatting.msg");

            foreach (var attachment in eml.Attachments)
            {
                if (!IsAttachmentInline(attachment))
                    continue;

                try { SaveAttachment(attachment, Data.Out/Guid.NewGuid().ToString()); }
                catch (Exception ex) { Console.WriteLine(ex.Message); }
            }
        }

        private static bool IsAttachmentInline(MapiAttachment attachment)
        {
            foreach (var property in attachment.ObjectData.Properties.Values)
            {
                if (property.Name != "\u0003ObjInfo")
                    continue;

                var odtPersist1 = BitConverter.ToUInt16(property.Data, 0);
                return (odtPersist1 & (1 << (7 - 1))) == 0;
            }
            return false;
        }

        private static void SaveAttachment(MapiAttachment attachment, string fileName)
        {
            foreach (var property in attachment.ObjectData.Properties.Values)
            {
                if (property.Name != "Package")
                    continue;

                using (var fs = new FileStream(fileName, FileMode.Create, FileAccess.Write))
                {
                    fs.Write(property.Data, 0, property.Data.Length);
                }
            }
        }
    }
}
