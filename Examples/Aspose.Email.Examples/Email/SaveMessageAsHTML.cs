// Demonstrates how to save an EML message as HTML with default options, and with
// custom options for embedding resources and writing complete email headers.

using System;
using System.IO;

namespace Aspose.Email.Examples.Email
{
    internal static class SaveMessageAsHtml
    {
        public static void Run()
        {
            var eml = MailMessage.Load(Data.Email/"Message.eml");

            var defaultPath = Data.Out/"SaveAsHTML_out.html";
            eml.Save(defaultPath, SaveOptions.DefaultHtml);

            // EmbedIntoHtml inlines images and other resources, so the result is a single
            // self-contained file rather than an HTML file plus a folder of resources.
            var options = new HtmlSaveOptions
            {
                ResourceRenderingMode = ResourceRenderingMode.EmbedIntoHtml,
                HtmlFormatOptions = HtmlFormatOptions.WriteHeader | HtmlFormatOptions.WriteCompleteEmailAddress
            };

            var embeddedPath = Data.Out/"SaveAsHTML1_out.html";
            eml.Save(embeddedPath, options);

            Console.WriteLine($"Default options:  {new FileInfo(defaultPath).Length,8:N0} bytes  {defaultPath}");
            Console.WriteLine($"Embedded + header:{new FileInfo(embeddedPath).Length,8:N0} bytes  {embeddedPath}");
        }
    }
}
