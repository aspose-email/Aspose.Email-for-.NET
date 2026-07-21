// Demonstrates how to sign an email message with a DKIM signature using a private
// RSA key, so the receiving server can verify the message really came from you.

using System;
using Aspose.Email.DKIM;

namespace Aspose.Email.Examples.Email
{
    internal static class SignEmailsWithDkim
    {
        public static void Run()
        {
            // The private key matching the DNS TXT record of the signing domain.
            var rsa = PemReader.GetPrivateKey(Data.Smtp/"key2.pem");

            // "test" is the selector, which tells the receiver which DNS record to look
            // up; only the listed headers are covered by the signature.
            var signInfo = new DKIMSignatureInfo("test", "yandex.ru");
            signInfo.Headers.Add("From");
            signInfo.Headers.Add("Subject");

            var mailMessage = new MailMessage("useremail@gmail.com", "test@gmail.com")
            {
                Subject = "Signed DKIM message text body",
                Body = "This is a text body signed DKIM message"
            };

            var signedMsg = mailMessage.DKIMSign(rsa, signInfo);

            var outputPath = Data.Out/"SignEmailsWithDkim_out.eml";
            signedMsg.Save(outputPath, SaveOptions.DefaultEml);

            Console.WriteLine($"Signed with selector '{signInfo.Selector}' for domain '{signInfo.Domain}'.");
            Console.WriteLine($"Saved to {outputPath}");

            // Sending needs a real server, so it only runs once one is configured in
            // clientsettings.json. The signed message above is complete either way.
            if (!ClientBuilder.IsSmtpConfigured)
            {
                Console.WriteLine("Set Smtp.HostName in clientsettings.json to also send this message.");
                return;
            }

            using (var client = ClientBuilder.Smtp(AuthType.Basic))
            {
                client.Send(signedMsg);
                Console.WriteLine("Message sent.");
            }
        }
    }
}
