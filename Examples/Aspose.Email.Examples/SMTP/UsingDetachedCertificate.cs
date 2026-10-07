// Demonstrates how to sign a message with a detached S/MIME signature.
//
// A detached signature (multipart/signed) leaves the message body readable as it is and
// adds the signature as a separate part, so mail programs without S/MIME support still
// show the text. An opaque signature wraps body and signature into one binary part that
// only S/MIME-aware programs can open. The second argument of AttachSignature chooses
// between the two.
//
// The signed message is saved to Out, and sent to your own address when SMTP is
// configured.

using System;
using System.Security.Cryptography.X509Certificates;

namespace Aspose.Email.Examples.SMTP
{
    internal static class UsingDetachedCertificate
    {
        public static void Run()
        {
            var certificate = new X509Certificate2(Data.Smtp/"MartinCertificate.pfx", "anothertestaccount");

            using (var client = ClientBuilder.IsSmtpConfigured ? ClientBuilder.Smtp(AuthType.Basic) : null)
            {
                var sender = client != null ? client.Username : "sender@example.com";
                var recipient = client != null ? client.Username : "receiver@example.com";

                var message = new MailMessage(sender, recipient,
                    "Signed message", "This text stays readable without S/MIME support.");

                var signed = message.AttachSignature(certificate, true);
                Console.WriteLine($"Signed with {certificate.Subject}, is signed = {signed.IsSigned}");

                var outputPath = Data.Out/"UsingDetachedCertificate_out.eml";
                signed.Save(outputPath, SaveOptions.DefaultEml);
                Console.WriteLine($"Saved to {outputPath}");

                if (client == null)
                {
                    Console.WriteLine("Set Smtp.HostName in clientsettings.json to also send it.");
                    return;
                }

                client.Send(signed);
                Console.WriteLine($"Sent to {recipient}.");
            }
        }
    }
}
