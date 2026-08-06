// Demonstrates how to attach a cloud link to a MIME message. A reference attachment
// carries the URL and the permission level, not the file itself.

using System;

namespace Aspose.Email.Examples.Email
{
    internal static class AddReferenceAttachmentToEml
    {
        public static void Run()
        {
            var eml = MailMessage.Load(Data.Email/"Message.eml");

            var refAttach = new ReferenceAttachment("https://example.com/shared/Document.docx")
            {
                Name = "Document.docx",
                ProviderType = AttachmentProviderType.OneDrivePro,
                PermissionType = AttachmentPermissionType.AnyoneCanEdit
            };

            eml.Attachments.Add(refAttach);

            Console.WriteLine($"Attachments: {eml.Attachments.Count}");
            foreach (var attachment in eml.Attachments)
                Console.WriteLine($"  {attachment.Name} ({attachment.ContentType.MediaType})");

            var outputPath = Data.Out/"AddReferenceAttachmentToEml_out.eml";
            eml.Save(outputPath, SaveOptions.DefaultEml);
            Console.WriteLine($"\nSaved to {outputPath}");
        }
    }
}
