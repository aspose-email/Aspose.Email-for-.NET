// Demonstrates how to take control of the TLS side of the connection: which protocol
// versions may be negotiated, and whether the server's certificate is trusted.
//
// By default the operating system validates the certificate. Passing a
// RemoteCertificateValidationCallback lets you log what the server presented, or pin a
// known certificate. Never simply return true from this callback in production - that
// switches off server authentication and invites man-in-the-middle attacks.

using System;
using System.Net.Security;
using Aspose.Email.Clients;
using Aspose.Email.Clients.Base;
using Aspose.Email.Clients.Smtp;

namespace Aspose.Email.Examples.SMTP
{
    internal static class ValidateSmtpServerCertificate
    {
        public static void Run()
        {
            RemoteCertificateValidationCallback validateCertificate = (sender, certificate, chain, errors) =>
            {
                Console.WriteLine($"Server certificate: {certificate?.Subject}");
                Console.WriteLine($"  issued by:     {certificate?.Issuer}");
                Console.WriteLine($"  expires:       {certificate?.GetExpirationDateString()}");
                Console.WriteLine($"  policy errors: {errors}");

                // Trust exactly what the operating system trusts.
                return errors == SslPolicyErrors.None;
            };

            using (var client = new SmtpClient("smtp.example.com", 587, "user@example.com", "password",
                validateCertificate))
            {
                client.SecurityOptions = SecurityOptions.SSLExplicit;

                // Refuse anything older than TLS 1.2. Versions the runtime does not
                // support are skipped silently.
                client.SupportedEncryption = EncryptionProtocols.Tls12 | EncryptionProtocols.Tls13;

                var valid = client.ValidateCredentials();
                Console.WriteLine($"\nConnected over TLS, credentials accepted: {valid}");
            }
        }
    }
}
