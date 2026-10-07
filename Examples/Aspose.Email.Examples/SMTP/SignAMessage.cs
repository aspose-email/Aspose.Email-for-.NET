// Demonstrates the S/MIME round trip: sign a message, encrypt it, then decrypt it and
// remove the signature again.
//
// Signing needs the sender's certificate with its private key (.pfx); encrypting needs
// the recipient's public certificate (.cer), and only the matching private key can
// decrypt. Here both belong to the same test certificate. UsingDetachedCertificate sends
// a signed message.

using System;
using System.Security.Cryptography.X509Certificates;

namespace Aspose.Email.Examples.SMTP
{
    internal static class SignAMessage
    {
        public static void Run()
        {
            var publicCert = new X509Certificate2(Data.Smtp/"MartinCertificate.cer");
            var privateCert = new X509Certificate2(Data.Smtp/"MartinCertificate.pfx", "anothertestaccount");

            var message = new MailMessage("sender@example.com", "receiver@example.com",
                "Signed and encrypted", "Test body of a signed message.");

            var signed = message.AttachSignature(privateCert);
            Console.WriteLine($"Signed:    is signed = {signed.IsSigned}");

            var encrypted = signed.Encrypt(publicCert);
            Console.WriteLine($"Encrypted: is encrypted = {encrypted.IsEncrypted}");
            encrypted.Save(Data.Out/"SignAMessage_encrypted_out.eml", SaveOptions.DefaultEml);

            var decrypted = encrypted.Decrypt(privateCert);
            Console.WriteLine($"Decrypted: is encrypted = {decrypted.IsEncrypted}, is signed = {decrypted.IsSigned}");

            var unsigned = decrypted.RemoveSignature();
            Console.WriteLine($"Unsigned:  is signed = {unsigned.IsSigned}, body = {unsigned.Body.Trim()}");

            unsigned.Save(Data.Out/"SignAMessage_out.eml", SaveOptions.DefaultEml);
            Console.WriteLine($"\nSaved the encrypted and the final message to {Data.Out}");
        }
    }
}
