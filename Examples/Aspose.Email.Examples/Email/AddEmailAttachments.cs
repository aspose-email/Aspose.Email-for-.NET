// Demonstrates how to add multiple attachments to a MailMessage and save it as EML.

using System;

namespace Aspose.Email.Examples.Email
{
    internal static class AddEmailAttachments
    {
        public static void Run()
        {
            var eml = new MailMessage
            {
                From = "sender@from.com",
                To = "receiver@to.com",
                Subject = "This is message",
                Body = "This is body"
            };

            eml.Attachments.Add(new Attachment(Data.Email/"1.txt"));
            eml.Attachments.Add(new Attachment(Data.Email/"1.jpg"));
            eml.Attachments.Add(new Attachment(Data.Email/"1.doc"));
            eml.Attachments.Add(new Attachment(Data.Email/"1.rar"));
            eml.Attachments.Add(new Attachment(Data.Email/"1.pdf"));

            var outputPath = Data.Out/"AddAttachments.eml";
            eml.Save(outputPath);

            foreach (var attachment in eml.Attachments)
                Console.WriteLine($"Attached: {attachment.Name}");

            Console.WriteLine($"\nSaved {eml.Attachments.Count} attachment(s) to {outputPath}");
        }
    }
}
