// Demonstrates how to replace the contents of an embedded MSG attachment with
// another embedded message.

using System;
using System.IO;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class ReplaceEmbeddedMSGAttachmentContents
    {
        public static void Run()
        {
            var message = MapiMessage.Load(Data.Mapi/"message3.msg");
            Console.WriteLine($"Attachment 1 before: {message.Attachments[1].DisplayName}");

            // Take the third attachment out as a standalone message...
            using (var ms = new MemoryStream())
            {
                message.Attachments[2].Save(ms);
                var replacement = MapiMessage.Load(ms);

                // ...and put it in place of the second one, under a new name.
                message.Attachments.Replace(1, "new 1", replacement);
            }

            Console.WriteLine($"Attachment 1 after:  {message.Attachments[1].DisplayName}");

            var outputPath = Data.Out/"ReplaceEmbeddedMSGAttachmentContents_out.msg";
            message.Save(outputPath);
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}
