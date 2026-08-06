// Demonstrates how to spot a winmail.dat attachment - a whole message packed into
// TNEF - among the ordinary attachments of an email.

using System;

namespace Aspose.Email.Examples.Email
{
    internal static class DetectTnefAttachment
    {
        public static void Run()
        {
            foreach (var fileName in new[] { "tnefWithMsgInside.eml", "Attachments.eml" })
            {
                // Loading with PreserveTnefAttachments keeps the winmail.dat part intact;
                // without it the library unpacks it into ordinary attachments.
                var eml = MailMessage.Load(Data.Email/fileName,
                    new EmlLoadOptions { PreserveTnefAttachments = true });

                Console.WriteLine($"{fileName}: {eml.Attachments.Count} attachment(s)");

                foreach (var attachment in eml.Attachments)
                    Console.WriteLine($"  {attachment.Name}: is TNEF? {attachment.IsTnef}");
            }
        }
    }
}
