// Demonstrates choosing the local IP address the client connects from.
//
// On a machine with several network interfaces or addresses, the BindIPEndPoint event
// lets you pick the local end of the connection - for example the address your SPF
// record allows, or an interface on a particular network. The handler receives the
// server's endpoint and returns the local one; port 0 lets the system pick the port.

using System;
using System.Net;
using System.Net.Sockets;

namespace Aspose.Email.Examples.SMTP
{
    internal static class SetSpecificIpAddress
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
                client.BindIPEndPoint += remoteEndPoint =>
                {
                    // Replace Any/IPv6Any with the local address to send from.
                    var localAddress = remoteEndPoint.AddressFamily == AddressFamily.InterNetworkV6
                        ? IPAddress.IPv6Any
                        : IPAddress.Any;

                    Console.WriteLine($"Connecting to {remoteEndPoint} from {localAddress}");
                    return new IPEndPoint(localAddress, 0);
                };

                var valid = client.ValidateCredentials();
                Console.WriteLine($"Connected and signed in: {valid}");
            }
        }
    }
}
