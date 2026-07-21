// Demonstrates how to change the sender, recipient, CC, and BCC addresses
// of a loaded EML message using friendly display names.

using System;

namespace Aspose.Email.Examples.Email
{
    internal static class ChangeEmailAddress
    {
        public static void Run()
        {
            var eml = MailMessage.Load(Data.Email/"test.eml");

            eml.From = new MailAddress("TimothyFairfield@from.com", "Timothy Fairfield");
            eml.To.Add(new MailAddress("kyle@to.com", "Kyle Huang"));
            eml.CC.Add(new MailAddress("guangzhou@cc.com", "Guangzhou Team"));
            eml.Bcc.Add(new MailAddress("ahaq@bcc.com", "Ammad ulHaq"));

            var outputPath = Data.Out/"MessageWithFriendlyName_out.eml";
            eml.Save(outputPath);

            // The display name is what a mail client shows instead of the raw address.
            Console.WriteLine($"From: {eml.From}");
            Console.WriteLine($"To:   {eml.To}");
            Console.WriteLine($"CC:   {eml.CC}");
            Console.WriteLine($"Bcc:  {eml.Bcc}");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}
