// Demonstrates how to permanently destroy all attachments in a saved MSG file.

using System;
using System.IO;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class DestroyAttachment
    {
        public static void Run()
        {
            // Work on a copy: DestroyAttachments rewrites the file it is given.
            var outputPath = Data.Out/"AttachmentsToDestroy_out.msg";
            MapiMessage.Load(Data.Mapi/"MsgWithAtt.msg").Save(outputPath);

            Console.WriteLine($"Attachments before: {MapiMessage.Load(outputPath).Attachments.Count}");
            Console.WriteLine($"File size before:   {new FileInfo(outputPath).Length,8:N0} bytes");

            // Unlike removing an attachment, this also wipes the bytes from the file, so
            // the content cannot be recovered from it afterwards.
            MapiMessage.DestroyAttachments(outputPath);

            Console.WriteLine($"Attachments after:  {MapiMessage.Load(outputPath).Attachments.Count}");
            Console.WriteLine($"File size after:    {new FileInfo(outputPath).Length,8:N0} bytes");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}
