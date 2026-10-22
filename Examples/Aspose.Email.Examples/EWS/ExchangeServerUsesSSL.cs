// Demonstrates connecting to Exchange over IMAP with implicit TLS.
//
// EWS always runs over HTTPS. For IMAP, Exchange listens on port 993 with implicit TLS
// (SecurityOptions.SSLImplicit): the connection is encrypted before the client signs
// in. The server's certificate is checked by the operating system - do not switch that
// off; IgnoringInvalidSSLCertificates shows how to trust a single test certificate.

using System;
using Aspose.Email.Clients;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.EWS
{
    internal static class ExchangeServerUsesSSL
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.SecurityOptions = SecurityOptions.SSLImplicit;

                client.SelectFolder(ImapFolderInfo.InBox);
                Console.WriteLine($"Connected to {client.Host}:{client.Port} using {client.SecurityOptions}; " +
                                  $"{client.CurrentFolder.TotalMessageCount} message(s) in the Inbox.");
            }
        }
    }
}
