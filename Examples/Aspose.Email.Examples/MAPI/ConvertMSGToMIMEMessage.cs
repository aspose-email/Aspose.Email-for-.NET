// Demonstrates how to convert a MapiMessage to a MIME MailMessage and save it as EML.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class ConvertMsgToMimeMessage
    {
        public static void Run()
        {
            var msg = new MapiMessage(
                "sender@test.com",
                "recipient1@test.com; recipient2@test.com",
                "Test Subject",
                "This is a body of message.");

            // ConvertAsTnef wraps the Outlook-specific properties in a TNEF part, so
            // they survive the trip through MIME.
            var options = new MailConversionOptions { ConvertAsTnef = true };
            var mail = msg.ToMailMessage(options);

            var outputPath = Data.Out/"ConvertMSGToMIMEMessage_out.eml";
            mail.Save(outputPath, SaveOptions.DefaultEml);

            Console.WriteLine($"Converted: {mail.Subject}");
            Console.WriteLine($"TNEF-encoded: {mail.OriginalIsTnef}");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}
