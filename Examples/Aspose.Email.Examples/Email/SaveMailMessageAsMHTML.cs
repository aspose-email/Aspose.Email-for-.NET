// Demonstrates how to load an EML file and save it in MHTML format.

using System;

namespace Aspose.Email.Examples.Email
{
    internal static class SaveMailMessageAsMhtml
    {
        public static void Run()
        {
            var eml = MailMessage.Load(Data.Email/"Message.eml");

            // MHTML packs the body and all its resources into a single file, which is
            // what makes it convenient for archiving a message.
            var outputPath = Data.Out/"AnEmail_out.mhtml";
            eml.Save(outputPath, SaveOptions.DefaultMhtml);

            Console.WriteLine($"Converted: {eml.Subject}");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}
