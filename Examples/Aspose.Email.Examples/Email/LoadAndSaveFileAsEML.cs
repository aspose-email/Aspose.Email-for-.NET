// Demonstrates how to load an existing EML file and save it back in EML format.

using System;

namespace Aspose.Email.Examples.Email
{
    internal static class LoadAndSaveFileAsEml
    {
        public static void Run()
        {
            var eml = MailMessage.Load(Data.Email/"Attachments.eml");

            var outputPath = Data.Out/"LoadAndSaveFileAsEML_out.eml";
            eml.Save(outputPath, SaveOptions.DefaultEml);

            Console.WriteLine($"Loaded: {eml.Subject} ({eml.Attachments.Count} attachment(s))");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}
