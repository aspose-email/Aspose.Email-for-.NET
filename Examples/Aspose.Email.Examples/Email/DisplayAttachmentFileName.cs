// Demonstrates how to load an EML message and display the file name of each attachment.

using System;

namespace Aspose.Email.Examples.Email
{
    internal static class DisplayAttachmentFileName
    {
        public static void Run()
        {
            var eml = MailMessage.Load(Data.Email/"Attachments.eml");

            foreach (var attachment in eml.Attachments)
                Console.WriteLine(attachment.Name);
        }
    }
}
