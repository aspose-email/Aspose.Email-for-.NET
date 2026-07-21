// Demonstrates how to update image resources and linked resources inside a TNEF EML file
// and save it with TNEF attachment preservation.

using System;
using System.IO;

namespace Aspose.Email.Examples.Email
{
    internal static class UpdateTnefAttachments
    {
        private static int updated;

        public static void Run()
        {
            updated = 0;

            var eml = MailMessage.Load(Data.Email/"tnefEml1.eml");
            UpdateResources(eml, Data.Email/"Untitled.jpg");

            // PreserveTnefAttachments keeps the updated resources inside the TNEF part
            // rather than re-encoding them as ordinary MIME attachments.
            var saveOptions = new EmlSaveOptions(MailMessageSaveType.EmlFormat)
            {
                FileCompatibilityMode = FileCompatibilityMode.PreserveTnefAttachments
            };

            var outputPath = Data.Out/"UpdateTNEFAttachments_out.eml";
            eml.Save(outputPath, saveOptions);

            Console.WriteLine($"Replaced {updated} resource(s).");
            Console.WriteLine($"Saved to {outputPath}");
        }

        // Recurses into embedded messages, so resources nested inside them are updated too.
        private static void UpdateResources(MailMessage msg, string imgFileName)
        {
            foreach (var attachment in msg.Attachments)
            {
                var mediaType = attachment.ContentType.MediaType;
                var name = attachment.ContentType.Name;

                if (mediaType == "image/png" ||
                    mediaType == "application/octet-stream" && Path.GetExtension(name) == ".jpg")
                {
                    attachment.ContentStream = new MemoryStream(File.ReadAllBytes(imgFileName));
                    updated++;
                }
                else if (mediaType == "message/rfc822" ||
                         mediaType == "application/octet-stream" && Path.GetExtension(name) == ".msg")
                {
                    var ms = new MemoryStream();
                    attachment.Save(ms);
                    ms.Position = 0;

                    var embeddedMessage = MailMessage.Load(ms);
                    UpdateResources(embeddedMessage, imgFileName);

                    var msProcessed = new MemoryStream();
                    embeddedMessage.Save(msProcessed, SaveOptions.DefaultMsgUnicode);
                    msProcessed.Position = 0;
                    attachment.ContentStream = msProcessed;
                }
            }

            foreach (var lr in msg.LinkedResources)
            {
                if (lr.ContentType.MediaType == "image/png")
                {
                    lr.ContentStream = new MemoryStream(File.ReadAllBytes(imgFileName));
                    updated++;
                }
            }
        }
    }
}
