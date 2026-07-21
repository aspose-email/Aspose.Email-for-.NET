// Demonstrates how to specify To, CC, and BCC recipient addresses on a MailMessage
// using collection initializer syntax.

using System;

namespace Aspose.Email.Examples.Email
{
    internal static class SpecifyRecipientAddresses
    {
        public static void Run()
        {
            // The address collections support collection initializers, so several
            // recipients can be given without repeating Add() calls.
            var eml = new MailMessage
            {
                From = "sender@sender.com",
                Subject = "Message with several recipients",
                To = { "receiver1@receiver.com", "receiver2@receiver.com", "receiver3@receiver.com" },
                CC = { "CC1@receiver.com", "CC2@receiver.com" },
                Bcc = { "Bcc1@receiver.com", "Bcc2@receiver.com" }
            };

            var outputPath = Data.Out/"SpecifyRecipientAddresses_out.eml";
            eml.Save(outputPath, SaveOptions.DefaultEml);

            Console.WriteLine($"To:  {eml.To.Count} recipient(s) - {eml.To}");
            Console.WriteLine($"CC:  {eml.CC.Count} recipient(s) - {eml.CC}");
            Console.WriteLine($"Bcc: {eml.Bcc.Count} recipient(s) - {eml.Bcc}");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}
