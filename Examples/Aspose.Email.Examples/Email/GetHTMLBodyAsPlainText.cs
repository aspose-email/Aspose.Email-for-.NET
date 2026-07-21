// Demonstrates how to extract the plain text representation of an HTML email body,
// with and without hyperlink URLs included.

using System;

namespace Aspose.Email.Examples.Email
{
    internal static class GetHtmlBodyAsPlainText
    {
        public static void Run()
        {
            var eml = MailMessage.Load(Data.Email/"HtmlWithUrlSample.eml");

            // Include hyperlink URLs in the plain text output
            var bodyWithUrl = eml.GetHtmlBodyText(true);
            // Exclude hyperlink URLs from the plain text output
            var bodyWithoutUrl = eml.GetHtmlBodyText(false);

            Console.WriteLine($"Body with URL: {bodyWithUrl}");
            Console.WriteLine($"Body without URL: {bodyWithoutUrl}");
        }
    }
}
