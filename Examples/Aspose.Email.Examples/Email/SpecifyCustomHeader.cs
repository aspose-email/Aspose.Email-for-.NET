// Demonstrates how to add a custom MIME header to a MailMessage.

using System;

namespace Aspose.Email.Examples.Email
{
    internal static class SpecifyCustomHeader
    {
        public static void Run()
        {
            var eml = new MailMessage
            {
                ReplyToList = "reply@reply.com",
                From = "sender@sender.com",
                To = "receiver1@receiver.com",
                Subject = "test mail"
            };

            // Any header name works - by convention custom ones are prefixed with "X-".
            eml.Headers.Add("secret-header", "mystery");

            var outputPath = Data.Out/"SpecifyCustomHeader_out.eml";
            eml.Save(outputPath, SaveOptions.DefaultEml);

            Console.WriteLine($"secret-header: {eml.Headers["secret-header"]}");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}
