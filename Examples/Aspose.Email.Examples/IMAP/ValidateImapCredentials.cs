// Demonstrates how to check a login before doing any real work.
//
// ValidateCredentials connects, authenticates and reports the outcome as a bool, which
// suits a "Test connection" button in a settings dialog. Network problems (wrong host,
// port blocked) still surface as exceptions, so they can be told apart from a
// rejected password.

using System;

namespace Aspose.Email.Examples.IMAP
{
    internal static class ValidateImapCredentials
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                try
                {
                    var valid = client.ValidateCredentials();

                    Console.WriteLine(valid
                        ? $"The server accepted the credentials of {client.Username}."
                        : $"The server rejected the credentials of {client.Username}.");
                    Console.WriteLine($"Connection state: {client.ConnectionState}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Could not reach {client.Host}:{client.Port}: {ex.Message}");
                }
            }
        }
    }
}
