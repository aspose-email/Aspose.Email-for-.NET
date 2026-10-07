// Demonstrates how to encrypt the connection to the IMAP server.
//
// There are two ways, tied to the port:
// - SSLImplicit (usually port 993): the connection is encrypted from the first byte.
// - SSLExplicit (STARTTLS, usually port 143): the client connects in plain text and
//   upgrades the connection before signing in.
// SecurityOptions.Auto lets the client work it out; setting the option explicitly
// removes the guesswork. Never sign in over SecurityOptions.None.

using System;
using Aspose.Email.Clients;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class SSLEnabledIMAPServer
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.SecurityOptions = client.Port == 993 ? SecurityOptions.SSLImplicit : SecurityOptions.SSLExplicit;

                client.SelectFolder(ImapFolderInfo.InBox);

                Console.WriteLine($"Connected to {client.Host}:{client.Port} using {client.SecurityOptions}.");
                Console.WriteLine($"The Inbox holds {client.CurrentFolder.TotalMessageCount} message(s).");
            }
        }
    }
}
