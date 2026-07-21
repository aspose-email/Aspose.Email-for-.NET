// Demonstrates how to render hyperlinks from an email's HTML body using custom formatters —
// one that includes the href URL and one that shows only the link text.

using System;

namespace Aspose.Email.Examples.Email
{
    internal static class CustomHyperlinkRendering
    {
        public static void Run()
        {
            var msg = MailMessage.Load(Data.Email/"LinksSample.eml");
            Console.WriteLine(msg.GetHtmlBodyText(RenderHyperlinkWithHref));
            Console.WriteLine(msg.GetHtmlBodyText(RenderHyperlinkWithoutHref));
        }

        private static string RenderHyperlinkWithHref(string source)
        {
            var start = source.IndexOf("href=\"", StringComparison.Ordinal) + "href=\"".Length;
            var end = source.IndexOf("\"", start + "href=\"".Length, StringComparison.Ordinal);
            var href = source.Substring(start, end - start);
            start = source.IndexOf(">", StringComparison.Ordinal) + 1;
            end = source.IndexOf("<", start, StringComparison.Ordinal);
            var text = source.Substring(start, end - start);
            return $"{text}<{href}>";
        }

        private static string RenderHyperlinkWithoutHref(string source)
        {
            var start = source.IndexOf(">", StringComparison.Ordinal) + 1;
            var end = source.IndexOf("<", start, StringComparison.Ordinal);
            return source.Substring(start, end - start);
        }
    }
}
