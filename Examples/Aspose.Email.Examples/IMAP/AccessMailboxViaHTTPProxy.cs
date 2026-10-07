// Demonstrates connecting to the mailbox through an HTTP proxy.
//
// Where outgoing connections must go through a proxy, assign an HttpProxy to the
// client's Proxy property before the first command; the client then tunnels the IMAP
// connection through it (HTTP CONNECT). The proxy must allow tunnelling to the IMAP
// port. Pass a user name and password to the HttpProxy constructor if it needs them.
//
// Replace the proxy address below with yours before running the example.

using System;
using Aspose.Email.Clients;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class AccessMailboxViaHTTPProxy
    {
        public static void Run()
        {
            const string proxyHost = "proxy.example.com";
            const int proxyPort = 8080;

            using (var proxy = new HttpProxy(proxyHost, proxyPort))
            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.Proxy = proxy;

                client.SelectFolder(ImapFolderInfo.InBox);
                Console.WriteLine($"Connected through the HTTP proxy {proxyHost}:{proxyPort}.");
                Console.WriteLine($"The Inbox holds {client.CurrentFolder.TotalMessageCount} message(s).");
            }
        }
    }
}
