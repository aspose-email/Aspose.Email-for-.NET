// Demonstrates how to substitute your own icon for each attachment when a message is
// converted to HTML, so the output can match the look of the surrounding application.

using System;

namespace Aspose.Email.Examples.Email
{
    internal static class RenderCustomAttachmentIcons
    {
        public static void Run()
        {
            var outputDir = Data.OutSub("AttachmentIcons");

            var options = new HtmlSaveOptions
            {
                // SubstituteFromFile is what makes the handler's PathToResourceFile
                // replace the attachment's own content.
                ResourceRenderingMode = ResourceRenderingMode.SubstituteFromFile,
                HtmlFormatOptions = HtmlFormatOptions.WriteHeader
            };
            options.ResourceHtmlRendering += SetAttachmentIcon;

            var mailMessage = MailMessage.Load(Data.Email/"Attachments.eml");

            var outputPath = outputDir/"RenderCustomAttachmentIcons_out.html";
            mailMessage.Save(outputPath, options);

            Console.WriteLine($"\nSaved to {outputPath}");
        }

        private static void SetAttachmentIcon(object sender, ResourceHtmlRenderingEventArgs e)
        {
            var attachment = sender as AttachmentBase;
            var name = attachment?.ContentType.Name ?? string.Empty;

            e.Caption = name;

            // Point at whichever icon file suits the attachment type. The sample data
            // has no icon set, so a stand-in image is used for every type here.
            if (name.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
                e.PathToResourceFile = Data.Email/"1.jpg";
            else if (name.EndsWith(".doc", StringComparison.OrdinalIgnoreCase))
                e.PathToResourceFile = Data.Email/"Untitled.jpg";
            else
                e.PathToResourceFile = Data.Email/"1.jpg";

            Console.WriteLine($"  {name} -> {e.PathToResourceFile}");
        }
    }
}
