// Demonstrates how to create a forwarded message from an existing MSG using ForwardMessageBuilder.

using System;
using Aspose.Email.Mapi;
using Aspose.Email.Tools;

namespace Aspose.Email.Examples.MAPI
{
    internal static class CreateForwardMessage
    {
        public static void Run()
        {
            var original = MapiMessage.Load(Data.Mapi/"message1.msg");

            var builder = new ForwardMessageBuilder
            {
                AdditionMode = OriginalMessageAdditionMode.Textpart
            };

            // Textpart quotes the original message in the body of the forward.
            var forward = builder.BuildResponse(original);

            var outputPath = Data.Out/"forward_out.msg";
            forward.Save(outputPath);

            Console.WriteLine($"Original: {original.Subject}");
            Console.WriteLine($"Forward:  {forward.Subject}");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}
