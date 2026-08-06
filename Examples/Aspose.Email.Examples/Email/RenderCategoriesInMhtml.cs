// Demonstrates how to add the message categories to an MHTML export and translate
// the label they are rendered under.

using System;
using System.IO;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.Email
{
    internal static class RenderCategoriesInMhtml
    {
        public static void Run()
        {
            var msg = new MapiMessage("from@example.com", "to@example.com", "subject", "body")
            {
                Categories = new[] { "Urgently", "Important" }
            };

            var saveOptions = new MhtSaveOptions
            {
                MhtFormatOptions = MhtFormatOptions.WriteHeader
            };

            // Categories are not rendered unless they are added to RenderingHeaders.
            saveOptions.FormatTemplates[MhtTemplateName.Categories] =
                saveOptions.FormatTemplates[MhtTemplateName.Categories].Replace("Categories", "Les catégories");
            saveOptions.RenderingHeaders.Add(MhtTemplateName.Categories);

            var outputPath = Data.Out/"RenderCategoriesInMhtml_out.mhtml";
            msg.Save(outputPath, saveOptions);

            Console.WriteLine($"Categories: {string.Join(", ", msg.Categories)}");

            foreach (var line in File.ReadLines(outputPath))
            {
                if (line.IndexOf("catégories", StringComparison.OrdinalIgnoreCase) >= 0)
                    Console.WriteLine($"  {line.Trim()}");
            }

            Console.WriteLine($"\nSaved to {outputPath}");
        }
    }
}
