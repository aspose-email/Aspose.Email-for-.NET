// Demonstrates sending through an HTTP proxy.
//
// Where outgoing connections must go through a proxy, assign an HttpProxy to the
// client's Proxy property; the client then tunnels the SMTP connection through it
// (HTTP CONNECT). The proxy must allow tunnelling to the SMTP port. Pass a user name and
// password to the HttpProxy constructor if the proxy requires them.
//
// Replace the proxy address below with yours before running the example.

using System;
using Aspose.Email.Clients;

namespace Aspose.Email.Examples.SMTP
{
    internal static class SendEmailViaHttpProxy
    {
        public static void Run()
        {
            if (!ClientBuilder.IsSmtpConfigured)
            {
                SmtpExampleInfo.PrintNotConfigured();
                return;
            }

            const string proxyHost = "proxy.example.com";
            const int proxyPort = 8080;

            using (var proxy = new HttpProxy(proxyHost, proxyPort))
            using (var client = ClientBuilder.Smtp(AuthType.Basic))
            {
                client.Proxy = proxy;

                client.Send(new MailMessage(client.Username, client.Username, "Sent through an HTTP proxy", "Body"));
                Console.WriteLine($"Sent to {client.Username} through the HTTP proxy {proxyHost}:{proxyPort}.");
            }
        }
    }
}
