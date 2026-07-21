// Demonstrates how to load an MSG file and save its embedded message attachments to disk.

using System;

namespace Aspose.Email.Examples.Email
{
    internal static class ExtractEmbeddedObjectsFromEmail
    {
        public static void Run()
        {
            var eml = MailMessage.Load(Data.Email/"A new opportunities.msg", new MsgLoadOptions());

            foreach (var attachment in eml.Attachments)
            {
                Console.WriteLine(attachment.Name);
                attachment.Save(Data.Out/attachment.Name);
            }
        }
    }
}
