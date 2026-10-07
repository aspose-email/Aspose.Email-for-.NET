// Demonstrates SMTP command pipelining (RFC 2920).
//
// Without pipelining the client waits for the server's reply to every command - MAIL
// FROM, each RCPT TO, DATA - before sending the next. With it, the client sends them in
// one go and reads the replies afterwards, which saves round trips when sending many
// messages over a slow link. PipeliningMode.Auto uses it only if the server announces
// PIPELINING; Enabled and Disabled force the choice.

using System;
using System.Diagnostics;
using System.Linq;
using Aspose.Email.Clients;

namespace Aspose.Email.Examples.SMTP
{
    internal static class EnableSmtpPipelining
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
                client.UsePipelining = PipeliningMode.Auto;

                var messages = Enumerable.Range(1, 5)
                    .Select(i => new MailMessage(client.Username, client.Username, $"Pipelined message {i}", "Body"))
                    .ToList();

                var watch = Stopwatch.StartNew();
                client.Send(messages);

                Console.WriteLine($"Sent {messages.Count} message(s) in {watch.ElapsedMilliseconds} ms.");
                Console.WriteLine($"  client mode:          {client.UsePipelining.ClientMode}");
                Console.WriteLine($"  supported by server:  {client.UsePipelining.SupportedByServer}");
                Console.WriteLine($"  pipelining was used:  {client.UsePipelining.PipeliningEnabled}");
            }
        }
    }
}
