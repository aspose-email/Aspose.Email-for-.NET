// Demonstrates how to check an SMTP login without sending anything.
//
// ValidateCredentials connects, authenticates and reports the outcome as a bool, which
// suits a "Test connection" button in a settings dialog. Network problems (wrong host,
// port blocked) still surface as exceptions, so they can be told apart from a rejected
// password.

using System;

namespace Aspose.Email.Examples.SMTP
{
    internal static class ValidateSmtpCredentials
    {
        public static void Run()
        {
            if (!ClientBuilder.IsSmtpConfigured)
            {
                SmtpExampleInfo.PrintNotConfigured();
                return;
            }

            using (var client = ClientBuilder.Smtp(AuthType.Basic))
            {
                try
                {
                    var valid = client.ValidateCredentials();

                    Console.WriteLine(valid
                        ? $"{client.Host} accepted the credentials of {client.Username}."
                        : $"{client.Host} rejected the credentials of {client.Username}.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Could not reach {client.Host}:{client.Port}: {ex.Message}");
                }
            }
        }
    }
}
