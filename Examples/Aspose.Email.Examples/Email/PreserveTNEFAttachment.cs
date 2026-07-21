// Demonstrates how to load an EML file and save it with TNEF attachment preservation.

using System;

namespace Aspose.Email.Examples.Email
{
    internal static class PreserveTnefAttachment
    {
        public static void Run()
        {
            var eml = MailMessage.Load(Data.Email/"PreserveOriginalBoundaries.eml");

            // Without PreserveTnefAttachments the TNEF part would be rewritten as plain
            // MIME attachments, which loses the Outlook-specific data it carries.
            var emlSaveOptions = new EmlSaveOptions(MailMessageSaveType.EmlFormat)
            {
                FileCompatibilityMode = FileCompatibilityMode.PreserveTnefAttachments
            };

            var outputPath = Data.Out/"PreserveTNEFAttachment_out.eml";
            eml.Save(outputPath, emlSaveOptions);

            Console.WriteLine($"Message:          {eml.Subject}");
            Console.WriteLine($"Originally TNEF:  {eml.OriginalIsTnef}");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}
