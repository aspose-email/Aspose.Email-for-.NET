// Demonstrates how to embed a MapiMessage as an attachment in another MapiMessage.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class EmbedMessageAsAttachment
    {
        public static void Run()
        {
            var message = new MapiMessage("from@test.com", "to@test.com", "Subj", "This is a message body");
            var attachMsg = MapiMessage.Load(Data.Mapi/"message.msg");

            // Passing a MapiMessage rather than bytes makes it a true embedded message,
            // which Outlook opens as a message instead of offering it as a file.
            message.Attachments.Add("Weekly report.msg", attachMsg);

            var outputPath = Data.Out/"WithEmbeddedMsg_out.msg";
            message.Save(outputPath);

            Console.WriteLine($"Outer message:    {message.Subject}");
            Console.WriteLine($"Embedded message: {attachMsg.Subject}");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}
