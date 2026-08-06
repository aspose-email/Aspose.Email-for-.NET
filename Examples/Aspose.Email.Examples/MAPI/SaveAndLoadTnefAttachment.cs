// Demonstrates how to move a single attachment through the TNEF (winmail.dat)
// representation: SaveToTnef writes it out, LoadFromTnef reads it back.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class SaveAndLoadTnefAttachment
    {
        public static void Run()
        {
            var msg = MapiMessage.Load(Data.Mapi/"MsgWithAtt.msg");
            Console.WriteLine($"Attachments in the source message: {msg.Attachments.Count}");

            var tnefPath = Data.Out/"winmail.dat";
            msg.Attachments[0].SaveToTnef(tnefPath);
            Console.WriteLine($"Attachment '{msg.Attachments[0].LongFileName}' saved as TNEF to {tnefPath}");

            // Reading it back produces an attachment that can be added to any message.
            var fromTnef = MapiAttachment.LoadFromTnef(tnefPath);
            msg.Attachments.Add(fromTnef);

            Console.WriteLine($"Attachments after adding it back: {msg.Attachments.Count}");
            foreach (var attachment in msg.Attachments)
                Console.WriteLine($"  {attachment.LongFileName}");

            var outputPath = Data.Out/"SaveAndLoadTnefAttachment_out.msg";
            msg.Save(outputPath);
            Console.WriteLine($"\nSaved to {outputPath}");
        }
    }
}
