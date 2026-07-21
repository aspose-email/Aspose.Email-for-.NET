// Demonstrates how to set an HTML body on a MailMessage and save it.

using System;

namespace Aspose.Email.Examples.Email
{
    internal static class SetHtmlBody
    {
        public static void Run()
        {
            // Setting HtmlBody also fills TextBody with a plain-text rendering, so
            // clients without HTML support still show something readable.
            var eml = new MailMessage
            {
                HtmlBody = "<html><body>This is the HTML body</body></html>"
            };

            var outputPath = Data.Out/"SetHTMLBody_out.eml";
            eml.Save(outputPath, SaveOptions.DefaultEml);

            Console.WriteLine($"HtmlBody: {eml.HtmlBody}");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}
