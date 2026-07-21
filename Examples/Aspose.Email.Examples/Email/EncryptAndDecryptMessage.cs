// Demonstrates how to encrypt a MailMessage with a public X.509 certificate
// and decrypt it with the corresponding private certificate.

using System;
using System.Security.Cryptography.X509Certificates;

namespace Aspose.Email.Examples.Email
{
    internal static class EncryptAndDecryptMessage
    {
        public static void Run()
        {
            var publicCert = new X509Certificate2(Data.Email/"MartinCertificate.cer");
            var privateCert = new X509Certificate2(Data.Email/"MartinCertificate.pfx", "anothertestaccount");

            var eml = new MailMessage
            {
                From = "atneostthaecrcount@gmail.com",
                To = "atneostthaecrcount@gmail.com",
                Subject = "Test subject",
                Body = "Test Body"
            };

            var encryptedEml = eml.Encrypt(publicCert);
            Console.WriteLine(encryptedEml.IsEncrypted ? "It's encrypted." : "It's NOT encrypted.");

            var decryptedEml = encryptedEml.Decrypt(privateCert);
            Console.WriteLine(decryptedEml.IsEncrypted ? "It's encrypted." : "It's NOT encrypted.");
        }
    }
}
