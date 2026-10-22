// Demonstrates accepting the self-signed certificate of one known test server - and
// nothing else.
//
// Test and lab Exchange servers often use self-signed certificates, which .NET rejects.
// Never switch validation off altogether: a callback that always returns true accepts
// any certificate and exposes credentials to man-in-the-middle attacks. Instead accept
// only the exact certificate you expect, identified by its thumbprint, and keep the
// operating system's checks for everything else. Put your server's thumbprint below.

using System;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;

namespace Aspose.Email.Examples.EWS
{
    internal static class IgnoringInvalidSSLCertificates
    {
        private const string TrustedThumbprint = "0123456789ABCDEF0123456789ABCDEF01234567";

        public static void Run()
        {
            ServicePointManager.ServerCertificateValidationCallback += ValidateCertificate;

            try
            {
                using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
                    Console.WriteLine($"Connected to {client.MailboxInfo.MailboxUri}");
            }
            finally
            {
                // Do not leave the relaxed check in place for the rest of the process.
                ServicePointManager.ServerCertificateValidationCallback -= ValidateCertificate;
            }
        }

        private static bool ValidateCertificate(object sender, X509Certificate certificate, X509Chain chain,
            SslPolicyErrors errors)
        {
            if (errors == SslPolicyErrors.None)
                return true;

            var thumbprint = new X509Certificate2(certificate).Thumbprint;
            var trusted = string.Equals(thumbprint, TrustedThumbprint, StringComparison.OrdinalIgnoreCase);

            Console.WriteLine($"Certificate {certificate.Subject} has errors ({errors}); " +
                              (trusted ? "accepted, it is the known test certificate." : "rejected."));
            return trusted;
        }
    }
}
