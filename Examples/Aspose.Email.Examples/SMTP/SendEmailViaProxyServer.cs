// Demonstrates sending through a SOCKS proxy.
//
// Assign a SocksProxy to the client's Proxy property and the SMTP connection is made
// through it. SOCKS 4 and 5 are supported; SOCKS 5 can authenticate with a user name and
// password (the four-argument constructor).
//
// Replace the proxy address below with yours before running the example.

using System;
using Aspose.Email.Clients;

namespace Aspose.Email.Examples.SMTP
{
    internal static class SendEmailViaProxyServer
    {
        public static void Run()
        {
            if (!ClientBuilder.IsSmtpConfigured)
            {
                SmtpExampleInfo.PrintNotConfigured();
                return;
            }

            const string proxyHost = "socks.example.com";
            const int proxyPort = 1080;

            using (var proxy = new SocksProxy(proxyHost, proxyPort, SocksVersion.SocksV5))
            using (var client = ClientBuilder.Smtp(AuthType.Basic))
            {
                client.Proxy = proxy;

                client.Send(new MailMessage(client.Username, client.Username, "Sent through a SOCKS proxy", "Body"));
                Console.WriteLine($"Sent to {client.Username} through the SOCKS 5 proxy {proxyHost}:{proxyPort}.");
            }
        }
    }
}
