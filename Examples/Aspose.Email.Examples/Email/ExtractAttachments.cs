// Demonstrates how to load an EML file, iterate its attachments,
// print each attachment name, and save each attachment to disk.

using System;

namespace Aspose.Email.Examples.Email
{
    internal static class ExtractAttachments
    {
        public static void Run()
        {
            var eml = MailMessage.Load(Data.Email/"Nothing will work.eml", new EmlLoadOptions());

            foreach (var attachment in eml.Attachments)
            {
                Console.WriteLine(attachment.Name);
                attachment.Save(Data.Out/$"ExtractAttachments_{attachment.Name}");
            }
        }
    }
}
