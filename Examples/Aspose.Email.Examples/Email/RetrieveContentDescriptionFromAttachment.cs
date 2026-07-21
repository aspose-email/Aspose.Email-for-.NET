// Demonstrates how to retrieve the Content-Description header value
// from an email attachment.

using System;

namespace Aspose.Email.Examples.Email
{
    internal static class RetrieveContentDescriptionFromAttachment
    {
        public static void Run()
        {
            var eml = MailMessage.Load(Data.Email/"EmailWithAttachEmbedded.eml");
            Console.WriteLine(eml.Attachments[0].Headers["Content-Description"]);
        }
    }
}
