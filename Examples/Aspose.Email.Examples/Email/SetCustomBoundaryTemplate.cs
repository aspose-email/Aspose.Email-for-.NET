// Demonstrates how to control the MIME boundary strings a saved EML uses. Some
// downstream systems expect a particular boundary shape, and {#} is the placeholder
// that gets replaced with the part number.

using System;
using System.IO;

namespace Aspose.Email.Examples.Email
{
    internal static class SetCustomBoundaryTemplate
    {
        public static void Run()
        {
            var eml = MailMessage.Load(Data.Email/"Attachments.eml");

            var saveOptions = new EmlSaveOptions(MailMessageSaveType.EmlFormat)
            {
                BoundariesTemplate = "boundary--{#}"
            };

            var outputPath = Data.Out/"SetCustomBoundaryTemplate_out.eml";
            eml.Save(outputPath, saveOptions);

            Console.WriteLine($"Saved with boundary template '{saveOptions.BoundariesTemplate}'");

            foreach (var line in File.ReadLines(outputPath))
            {
                if (line.IndexOf("boundary", StringComparison.OrdinalIgnoreCase) >= 0)
                    Console.WriteLine($"  {line.Trim()}");
            }

            Console.WriteLine($"\nSaved to {outputPath}");
        }
    }
}
