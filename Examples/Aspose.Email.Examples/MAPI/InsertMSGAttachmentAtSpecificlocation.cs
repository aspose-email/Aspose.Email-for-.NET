// Demonstrates how to insert a MapiMessage attachment at a specific index in an
// existing MSG file, rather than appending it at the end.

using System;
using System.IO;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class InsertMsgAttachmentAtSpecificLocation
    {
        public static void Run()
        {
            var message = MapiMessage.Load(Data.Mapi/"message3.msg");
            Console.WriteLine($"Attachments before: {message.Attachments.Count}");

            using (var ms = new MemoryStream())
            {
                // Take the third attachment out as a standalone message...
                message.Attachments[2].Save(ms);
                ms.Position = 0;
                var attachMsg = MapiMessage.Load(ms);

                // ...and insert a copy of it at index 1, shifting the rest along.
                message.Attachments.Insert(1, "new 11", attachMsg);
            }

            Console.WriteLine($"Attachments after:  {message.Attachments.Count}");
            Console.WriteLine($"Attachment at index 1: {message.Attachments[1].DisplayName}");

            var outputPath = Data.Out/"AttachmentAtSpecificLocation_out.msg";
            message.Save(outputPath);
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}
