// Demonstrates how to load an MSG file with TNEF attachment preservation enabled
// and list the names of the preserved attachments.

using System;

namespace Aspose.Email.Examples.Email
{
    internal static class ReadMessageByPreservingTnefAttachments
    {
        public static void Run()
        {
            var options = new MsgLoadOptions { PreserveTnefAttachments = true };
            var eml = MailMessage.Load(Data.Email/"EmbeddedImage1.msg", options);

            foreach (var attachment in eml.Attachments)
                Console.WriteLine(attachment.Name);
        }
    }
}
