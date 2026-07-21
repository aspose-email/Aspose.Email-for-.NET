// Demonstrates how to determine whether an attachment in an EML message
// is itself an embedded message.

using System;

namespace Aspose.Email.Examples.Email
{
    internal static class DetermineAttachmentEmbeddedMessage
    {
        public static void Run()
        {
            var eml = MailMessage.Load(Data.Email/"EmailWithAttachEmbedded.eml");

            Console.WriteLine(eml.Attachments[0].IsEmbeddedMessage
                ? "Attachment is an embedded message."
                : "Attachment is not an embedded message.");
        }
    }
}
