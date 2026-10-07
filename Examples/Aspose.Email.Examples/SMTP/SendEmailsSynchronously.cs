// Demonstrates synchronous sending: each Send call blocks until the server has accepted
// the message, so the messages go out strictly one after another and any failure is
// thrown right at the call that caused it.
//
// The client keeps its connection open between calls, so only the first Send pays for
// connecting and signing in. For non-blocking sending see SendEmailAsynchronously.

using System;
using System.Diagnostics;

namespace Aspose.Email.Examples.SMTP
{
    internal static class SendEmailsSynchronously
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
                for (var i = 1; i <= 3; i++)
                {
                    var watch = Stopwatch.StartNew();
                    client.Send(new MailMessage(client.Username, client.Username, $"Synchronous message {i}", "Body"));
                    Console.WriteLine($"Message {i} accepted after {watch.ElapsedMilliseconds} ms");
                }
            }
        }
    }
}
