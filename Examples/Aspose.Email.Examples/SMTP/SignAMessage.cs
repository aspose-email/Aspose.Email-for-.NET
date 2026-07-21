using System.Security.Cryptography.X509Certificates;
using Aspose.Email.Mime;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.SMTP
{
    class SignAMessage
    {
        public static void Run()
        {
            string publicCertFile = Data.Smtp + "MartinCertificate.cer";
            string privateCertFile = Data.Smtp + "MartinCertificate.pfx";
            X509Certificate2 publicCert = new X509Certificate2(publicCertFile);
            X509Certificate2 privateCert = new X509Certificate2(privateCertFile, "password");
            MailMessage msg = new MailMessage("userfrom@gmail.com", "userto@gmail.com", "Signed message only", "Test Body of signed message");
            MailMessage signed = msg.AttachSignature(privateCert);
            MailMessage encrypted = signed.Encrypt(publicCert);
            MailMessage decrypted = encrypted.Decrypt(privateCert);
            MailMessage unsigned = decrypted.RemoveSignature();//The original message with proper body
            MapiMessage mapi = MapiMessage.FromMailMessage(unsigned);
        }
    }
}