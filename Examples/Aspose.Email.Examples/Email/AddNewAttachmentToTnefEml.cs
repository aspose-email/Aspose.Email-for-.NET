// Demonstrates how to add a new attachment to a TNEF-encoded EML file and save it
// with TNEF attachment preservation.

using System;
using System.IO;

namespace Aspose.Email.Examples.Email
{
    internal static class AddNewAttachmentToTnefEml
    {
        public static void Run()
        {
            var tnefEml = MailMessage.Load(Data.Email/"tnefEml1.eml");
            Console.WriteLine($"Attachments before: {tnefEml.Attachments.Count}");

            using (var fileStream = File.OpenRead(Data.Email/"Untitled.jpg"))
            {
                tnefEml.Attachments.Add(new Attachment(fileStream, "Image.jpg", "image/jpg"));

                // Without PreserveTnefAttachments the TNEF part would be rewritten as a
                // plain MIME attachment, losing the original encoding.
                var saveOptions = new EmlSaveOptions(MailMessageSaveType.EmlFormat)
                {
                    FileCompatibilityMode = FileCompatibilityMode.PreserveTnefAttachments
                };

                var outputPath = Data.Out/"test_out.eml";
                tnefEml.Save(outputPath, saveOptions);

                Console.WriteLine($"Attachments after:  {tnefEml.Attachments.Count}");
                Console.WriteLine($"Saved to {outputPath}");
            }
        }
    }
}
