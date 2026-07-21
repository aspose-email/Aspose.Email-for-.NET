// Demonstrates how to create a reply-all message from an existing MSG using ReplyMessageBuilder.

using System;
using Aspose.Email.Mapi;
using Aspose.Email.Tools;

namespace Aspose.Email.Examples.MAPI
{
    internal static class CreateReplyMessage
    {
        public static void Run()
        {
            var original = MapiMessage.Load(Data.Mapi/"message1.msg");

            var builder = new ReplyMessageBuilder
            {
                ReplyAll = true,
                AdditionMode = OriginalMessageAdditionMode.Textpart,
                ResponseText = "<p><b>Dear Friend,</b></p> I want to do is introduce my co-author and co-teacher. " +
                    "<p><a href=\"www.google.com\">This is a first link</a></p>" +
                    "<p><a href=\"www.google.com\">This is a second link</a></p>"
            };

            // ReplyAll puts every original recipient on the reply, not just the sender;
            // Textpart quotes the original message below the response text.
            var reply = builder.BuildResponse(original);

            var outputPath = Data.Out/"reply_out.msg";
            reply.Save(outputPath);

            Console.WriteLine($"Original: {original.Subject} ({original.Recipients.Count} recipient(s))");
            Console.WriteLine($"Reply:    {reply.Subject} ({reply.Recipients.Count} recipient(s))");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}
