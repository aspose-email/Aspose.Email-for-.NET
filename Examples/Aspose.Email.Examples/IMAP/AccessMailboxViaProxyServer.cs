// Demonstrates connecting to the mailbox through a SOCKS proxy.
//
// Assign a SocksProxy to the client's Proxy property before the first command and the
// IMAP connection is made through it. SOCKS 4 and 5 are supported; SOCKS 5 can
// authenticate with a user name and password (the four-argument constructor).
//
// Replace the proxy address below with yours before running the example.

using System;
using Aspose.Email.Clients;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class AccessMailboxViaProxyServer
    {
        public static void Run()
        {
            const string proxyHost = "socks.example.com";
            const int proxyPort = 1080;

            using (var proxy = new SocksProxy(proxyHost, proxyPort, SocksVersion.SocksV5))
            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.Proxy = proxy;

                client.SelectFolder(ImapFolderInfo.InBox);
                Console.WriteLine($"Connected through the SOCKS 5 proxy {proxyHost}:{proxyPort}.");
                Console.WriteLine($"The Inbox holds {client.CurrentFolder.TotalMessageCount} message(s).");
            }
        }
    }
}
