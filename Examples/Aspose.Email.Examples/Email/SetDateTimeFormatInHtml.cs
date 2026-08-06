// Demonstrates how to control how the sent date is written into an HTML export,
// by overriding the DateTime format template.

using System;
using System.IO;

namespace Aspose.Email.Examples.Email
{
    internal static class SetDateTimeFormatInHtml
    {
        public static void Run()
        {
            var msg = MailMessage.Load(Data.Email/"Message.eml");
            Console.WriteLine($"Message date: {msg.Date}");

            Save(msg, "default", null);
            Save(msg, "custom", "MM d yyyy HH:mm tt");
        }

        private static void Save(MailMessage msg, string label, string dateTimeFormat)
        {
            var options = new HtmlSaveOptions
            {
                HtmlFormatOptions = HtmlFormatOptions.WriteHeader
            };

            if (dateTimeFormat != null)
                options.FormatTemplates[MhtTemplateName.DateTime] = dateTimeFormat;

            var outputPath = Data.Out/$"SetDateTimeFormatInHtml_{label}.html";
            msg.Save(outputPath, options);

            Console.WriteLine($"\n{label} ({dateTimeFormat ?? "library default"}) -> {outputPath}");

            foreach (var line in File.ReadLines(outputPath))
            {
                if (line.IndexOf("Sent:", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    Console.WriteLine($"  {line.Trim()}");
                    return;
                }
            }
        }
    }
}
