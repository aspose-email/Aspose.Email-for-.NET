// Demonstrates how to list attachments from a signed email, remove the signature,
// and list the attachments again from the unsigned message.

using System;

namespace Aspose.Email.Examples.Email
{
    internal static class RetrievingAttachmentsFromSignedEmail
    {
        public static void Run()
        {
            var signedEml = MailMessage.Load(Data.Email/"signed.eml");

            if (!signedEml.IsSigned)
                return;

            for (var i = 0; i < signedEml.Attachments.Count; i++)
                Console.WriteLine($"Signed email attachment {i}: {signedEml.Attachments[i].Name}");

            var eml = signedEml.RemoveSignature();
            Console.WriteLine("Signature removed.");

            for (var i = 0; i < eml.Attachments.Count; i++)
                Console.WriteLine($"Email attachment {i}: {eml.Attachments[i].Name}");
        }
    }
}
