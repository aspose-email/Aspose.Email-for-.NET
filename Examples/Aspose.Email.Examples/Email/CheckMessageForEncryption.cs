// Demonstrates how to check, encrypt, and decrypt an email message using X.509 certificates,
// verifying the IsEncrypted flag at each step.

using System;
using System.Security.Cryptography.X509Certificates;

namespace Aspose.Email.Examples.Email
{
    internal static class CheckMessageForEncryption
    {
        public static void Run()
        {
            var emlOrig = MailMessage.Load(Data.Email/"Message.msg", new MsgLoadOptions());
            Console.WriteLine($"Message is encrypted: {emlOrig.IsEncrypted}");

            var publicCert = new X509Certificate2(Data.Email/"MartinCertificate.cer");
            var privateCert = new X509Certificate2(Data.Email/"MartinCertificate.pfx", "anothertestaccount");

            var eml = emlOrig.Encrypt(publicCert);
            Console.WriteLine($"Message is encrypted: {eml.IsEncrypted}");

            eml = eml.Decrypt(privateCert);
            Console.WriteLine($"Message is encrypted: {eml.IsEncrypted}");
        }
    }
}
