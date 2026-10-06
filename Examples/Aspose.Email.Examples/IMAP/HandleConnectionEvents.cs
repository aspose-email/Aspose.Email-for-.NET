// Demonstrates the two connection-level hooks of an email client.
//
// OnConnect fires every time the client opens a connection to the server - useful for
// logging, or for noticing that a dropped connection was re-established.
// BindIPEndPoint lets you choose the local address the socket is bound to, which matters
// on machines with several network interfaces when mail must leave through a specific
// one. Port 0 lets the operating system pick the local port.

using System;
using System.Net;
using System.Net.Sockets;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class HandleConnectionEvents
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.BindIPEndPoint += remoteEndPoint =>
                {
                    // Replace Any/IPv6Any with the address of the interface to use.
                    var localAddress = remoteEndPoint.AddressFamily == AddressFamily.InterNetworkV6
                        ? IPAddress.IPv6Any
                        : IPAddress.Any;

                    Console.WriteLine($"{DateTime.Now:T}  binding to {localAddress} to reach {remoteEndPoint}");
                    return new IPEndPoint(localAddress, 0);
                };

                client.OnConnect += (sender, args) =>
                    Console.WriteLine($"{DateTime.Now:T}  connected to {client.Host}:{client.Port}");

                // The first command opens the connection and raises both events.
                client.SelectFolder(ImapFolderInfo.InBox);
                Console.WriteLine($"{DateTime.Now:T}  Inbox holds {client.CurrentFolder.TotalMessageCount} message(s)");
            }
        }
    }
}
