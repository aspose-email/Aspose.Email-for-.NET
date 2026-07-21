// Demonstrates how to extract and save both regular attachments and inline linked
// resources (embedded images) from an EML message.

using System;

namespace Aspose.Email.Examples.Email
{
    internal static class ExtractEmbeddedObjects
    {
        public static void Run()
        {
            var eml = MailMessage.Load(Data.Email/"EmailWithAttachEmbedded.eml");

            foreach (var attachment in eml.Attachments)
            {
                Console.WriteLine(attachment.Name);
                attachment.Save(Data.Out/attachment.Name);
            }

            foreach (var lr in eml.LinkedResources)
            {
                Console.WriteLine(lr.ContentType.Name);
                lr.Save(Data.Out/lr.ContentType.Name);
            }
        }
    }
}
