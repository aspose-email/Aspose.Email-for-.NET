// Demonstrates how to set standard email headers (ReplyTo, From, To, CC, BCC, Subject,
// Date) and a custom header on a MailMessage.

using System;

namespace Aspose.Email.Examples.Email
{
    internal static class SetEmailHeaders
    {
        public static void Run()
        {
            var eml = new MailMessage
            {
                ReplyToList = "reply@reply.com",
                From = "sender@sender.com",
                To = "receiver1@receiver.com",
                CC = "receiver2@receiver.com",
                Bcc = "receiver3@receiver.com",
                Subject = "test mail",
                Date = new DateTime(2006, 3, 6),
                XMailer = "Aspose.Email"
            };

            // Anything not covered by a property can be added as a raw header.
            eml.Headers.Add("secret-header", "mystery");

            var outputPath = Data.Out/"SetEmailHeaders_out.msg";
            eml.Save(outputPath, SaveOptions.DefaultMsg);

            foreach (var name in eml.Headers.AllKeys)
                Console.WriteLine($"{name}: {eml.Headers[name]}");

            Console.WriteLine($"\nSaved to {outputPath}");
        }
    }
}
