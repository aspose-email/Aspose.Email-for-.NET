// Demonstrates how to load an EML file and display the sender, recipients,
// subject, HTML body, and plain text body.

using System;

namespace Aspose.Email.Examples.Email
{
    internal static class DisplayEmailInformation
    {
        public static void Run()
        {
            var eml = MailMessage.Load(Data.Email/"test.eml");

            Console.WriteLine($"From: {eml.From}");
            Console.WriteLine($"To: {eml.To}");
            Console.WriteLine($"Subject: {eml.Subject}");
            Console.WriteLine($"HtmlBody: {eml.HtmlBody}");
            Console.WriteLine($"TextBody: {eml.Body}");
        }
    }
}
