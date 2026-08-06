// Demonstrates how to verify the S/MIME signature of a message and read the
// certificates it was signed with.

using System;
using System.Security.Cryptography.X509Certificates;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.Email
{
    internal static class CheckMessageSignature
    {
        public static void Run()
        {
            var manager = new SecureEmailManager();

            var eml = MailMessage.Load(Data.Email/"signed.eml");
            Report("signed.eml (MailMessage)", manager.CheckSignature(eml));

            // Sign a MapiMessage on the fly, so the MapiMessage overload has a subject too.
            var privateCert = new X509Certificate2(Data.Email/"MartinCertificate.pfx", "anothertestaccount");
            var msg = manager.AttachSignature(MapiMessage.Load(Data.Email/"Message.msg"), privateCert);
            Report("freshly signed MSG", manager.CheckSignature(msg));

            // Passing the certificate explicitly lets a message be verified even when
            // the signer is not in the machine's certificate store.
            Report("freshly signed MSG, explicit certificate",
                manager.CheckSignature(msg, privateCert));
        }

        private static void Report(string label, SmimeResult result)
        {
            Console.WriteLine($"{label}: {(result.IsSuccess ? "signature is valid" : "signature is not valid")}");

            if (result.Error != null)
                Console.WriteLine($"  error: {result.Error.Message}");

            foreach (var certificate in result.SigningCertificates)
                Console.WriteLine($"  signed by: {certificate.Subject}");
        }
    }
}
