// Demonstrates how to set a preferred text encoding on a MailMessage so that
// addresses, subject, and body are all encoded consistently.

using System;
using System.Text;

namespace Aspose.Email.Examples.Email
{
    internal static class SetDefaultTextEncoding
    {
        public static void Run()
        {
            // 28591 is Latin-1. Setting PreferredTextEncoding once applies it to the
            // addresses, the subject and the body, instead of setting each separately.
            var eml = new MailMessage
            {
                PreferredTextEncoding = Encoding.GetEncoding(28591),
                From = new MailAddress("dmo@domain.com", "démo"),
                To = new MailAddress("dmo@domain.com", "démo"),
                Subject = "démo",
                HtmlBody = "démo"
            };

            var outputPath = Data.Out/"SetDefaultTextEncoding_out.msg";
            eml.Save(outputPath, SaveOptions.DefaultMsg);

            Console.WriteLine($"Preferred encoding: {eml.PreferredTextEncoding.WebName}");
            Console.WriteLine($"Subject: {eml.Subject}");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}
