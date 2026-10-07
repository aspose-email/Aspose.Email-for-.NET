// Demonstrates sending a large batch over several connections at once.
//
// With UseMultiConnection enabled, Send(messages) spreads the batch over up to
// ConnectionsQuantity parallel connections. That can shorten big mailings, but servers
// limit connections and messages per account, so more connections are not always
// faster - measure with your server. SendOverDedicatedConnections shows how to manage
// the connections yourself.

using System;
using System.Diagnostics;
using System.Linq;
using Aspose.Email.Clients;

namespace Aspose.Email.Examples.SMTP
{
    internal static class SendWithMultiConnection
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
                client.UseMultiConnection = MultiConnectionMode.Enable;
                client.ConnectionsQuantity = 3;

                var self = client.Username;
                var messages = Enumerable.Range(1, 9)
                    .Select(i => new MailMessage(self, self, $"Multi-connection message {i}", "Body"))
                    .ToList();

                var watch = Stopwatch.StartNew();
                client.Send(messages);

                Console.WriteLine($"Sent {messages.Count} message(s) over up to {client.ConnectionsQuantity} " +
                                  $"connections in {watch.ElapsedMilliseconds} ms.");
            }
        }
    }
}
