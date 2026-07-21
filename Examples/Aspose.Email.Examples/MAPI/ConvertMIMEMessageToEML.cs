// Demonstrates how to load an Outlook MSG file and save it as an EML file.

using System;

namespace Aspose.Email.Examples.MAPI
{
    internal static class ConvertMimeMessageToEml
    {
        public static void Run()
        {
            // MailMessage.Load detects the format, so an MSG can be loaded without
            // naming the format explicitly.
            var msg = MailMessage.Load(Data.Mapi/"Message2.msg");

            var outputPath = Data.Out/"ConvertMIMEMessageToEML_out.eml";
            msg.Save(outputPath);

            Console.WriteLine($"Converted: {msg.Subject}");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}
