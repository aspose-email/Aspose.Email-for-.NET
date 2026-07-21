// Demonstrates how to strip the attachments out of an MSG file in place, without
// loading the whole message into memory.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class RemoveAttachmentsFromFile
    {
        public static void Run()
        {
            // Work on a copy: RemoveAttachments rewrites the file it is given.
            var outputPath = Data.Out/"AttachmentsToRemove_out.msg";
            MapiMessage.Load(Data.Mapi/"MsgWithAtt.msg").Save(outputPath);

            Console.WriteLine($"Attachments before: {MapiMessage.Load(outputPath).Attachments.Count}");

            // This static overload works directly on the file.
            MapiMessage.RemoveAttachments(outputPath);

            Console.WriteLine($"Attachments after:  {MapiMessage.Load(outputPath).Attachments.Count}");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}
