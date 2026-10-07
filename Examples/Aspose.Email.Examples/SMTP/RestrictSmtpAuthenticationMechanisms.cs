// Demonstrates how to see which SASL mechanisms the server offers and how to limit the
// ones the client may use.
//
// SupportedAuthentication is filled in from the server's EHLO response once the client
// has connected. AllowedAuthentication is the client-side filter, set before
// connecting: the client only signs in with a mechanism you allowed, which is how you
// keep it from ever falling back to, say, LOGIN or PLAIN.

using System;
using Aspose.Email.Clients;
using Aspose.Email.Clients.Smtp;

namespace Aspose.Email.Examples.SMTP
{
    internal static class RestrictSmtpAuthenticationMechanisms
    {
        public static void Run()
        {
            using (var client = new SmtpClient("smtp.example.com", 587, "user@example.com", "password",
                SecurityOptions.SSLExplicit))
            {
                // Challenge-response mechanisms only: the password itself is never sent.
                client.AllowedAuthentication =
                    SmtpKnownAuthenticationType.ScramSha256 |
                    SmtpKnownAuthenticationType.ScramSha1 |
                    SmtpKnownAuthenticationType.CramMD5;

                Console.WriteLine($"Allowed by the client:  {client.AllowedAuthentication}");

                try
                {
                    var loggedIn = client.ValidateCredentials();
                    Console.WriteLine(loggedIn
                        ? "Signed in with one of the allowed mechanisms."
                        : "Sign-in failed: wrong credentials, or no allowed mechanism is offered.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Sign-in failed: {ex.Message}");
                }

                Console.WriteLine($"Offered by the server:  {client.SupportedAuthentication}");
            }
        }
    }
}
