// Demonstrates how to add a plain text alternate view to a MailMessage, providing a
// fallback for email clients that do not support HTML.

using System;

namespace Aspose.Email.Examples.Email
{
    internal static class SetAlternateText
    {
        public static void Run()
        {
            var eml = new MailMessage
            {
                From = "sender@sender.com",
                To = "receiver@receiver.com",
                Subject = "Message with an alternate view",
                HtmlBody = "<html><body>This is the HTML body</body></html>"
            };

            // An alternate view carries the same content in a different format; the
            // client picks the best one it can render.
            eml.AlternateViews.Add(AlternateView.CreateAlternateViewFromString("Alternate Text"));

            var outputPath = Data.Out/"SetAlternateText_out.eml";
            eml.Save(outputPath, SaveOptions.DefaultEml);

            Console.WriteLine($"Alternate views: {eml.AlternateViews.Count}");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}
