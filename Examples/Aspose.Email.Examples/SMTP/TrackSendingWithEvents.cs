// Demonstrates how to follow the outcome of every message in a batch.
//
// SucceededSending fires for each message the server accepted, FailedSending for each
// one it did not, with the exception that explains why. The events may be raised on
// other threads - and in multi-connection mode, from several connections at once - so
// the handlers below only use thread-safe counters and a lock around the console.

using System;
using System.Collections.Generic;
using System.Threading;

namespace Aspose.Email.Examples.SMTP
{
    internal static class TrackSendingWithEvents
    {
        public static void Run()
        {
            if (!ClientBuilder.IsSmtpConfigured)
            {
                SmtpExampleInfo.PrintNotConfigured();
                return;
            }

            var consoleLock = new object();
            var succeeded = 0;
            var failed = 0;

            using (var client = ClientBuilder.Smtp(AuthType.Basic))
            {
                client.SucceededSending += (sender, e) =>
                {
                    Interlocked.Increment(ref succeeded);
                    lock (consoleLock)
                        Console.WriteLine($"  sent:   {e.Message.Subject}");
                };

                client.FailedSending += (sender, e) =>
                {
                    Interlocked.Increment(ref failed);
                    lock (consoleLock)
                        Console.WriteLine($"  failed: {e.Message.Subject} - {e.OperationError?.Message}");
                };

                var messages = new List<MailMessage>();
                for (var i = 1; i <= 3; i++)
                    messages.Add(new MailMessage(client.Username, client.Username, $"Weekly digest, part {i}", "Body"));

                Console.WriteLine($"Sending {messages.Count} message(s) to {client.Username}:");

                try
                {
                    client.Send(messages);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Send reported an error: {ex.Message}");
                }
            }

            Console.WriteLine($"\n{succeeded} succeeded, {failed} failed.");
        }
    }
}
