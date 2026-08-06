// Demonstrates how to change the name Outlook shows for an attachment. DisplayName
// is what the reader sees; the underlying file name stays untouched.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class RenameMessageAttachment
    {
        public static void Run()
        {
            var msg = MapiMessage.Load(Data.Mapi/"MsgWithAtt.msg");

            for (var i = 0; i < msg.Attachments.Count; i++)
            {
                var attachment = msg.Attachments[i];
                Console.WriteLine($"Before: DisplayName='{attachment.DisplayName}' LongFileName='{attachment.LongFileName}'");

                attachment.DisplayName = $"New display name {i + 1}";
            }

            var outputPath = Data.Out/"RenameMessageAttachment_out.msg";
            msg.Save(outputPath);

            var reloaded = MapiMessage.Load(outputPath);
            foreach (var attachment in reloaded.Attachments)
                Console.WriteLine($"After:  DisplayName='{attachment.DisplayName}' LongFileName='{attachment.LongFileName}'");

            Console.WriteLine($"\nSaved to {outputPath}");
        }
    }
}
