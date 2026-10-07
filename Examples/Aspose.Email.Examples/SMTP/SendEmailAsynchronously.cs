// Demonstrates sending without blocking the caller.
//
// SendAsync returns a Task, so the program can keep working - here it just reports
// progress - while the message is transferred, and await the result when it needs it.
// The CancellationToken abandons the send if it takes too long. SendMessagesAsync shows
// the parameter-object form of the same API.

using System;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;

namespace Aspose.Email.Examples.SMTP
{
    internal static class SendEmailAsynchronously
    {
        public static void Run()
        {
            RunAsync().GetAwaiter().GetResult();
        }

        private static async Task RunAsync()
        {
            if (!ClientBuilder.IsSmtpConfigured)
            {
                SmtpExampleInfo.PrintNotConfigured();
                return;
            }

            using (var client = ClientBuilder.Smtp(AuthType.Basic))
            using (var cancellation = new CancellationTokenSource(TimeSpan.FromSeconds(30)))
            {
                var message = new MailMessage(client.Username, client.Username, "Sent asynchronously", "Body");
                var watch = Stopwatch.StartNew();

                try
                {
                    var sending = client.SendAsync(message, cancellation.Token);
                    Console.WriteLine($"{watch.ElapsedMilliseconds,6} ms  sending started");

                    while (!sending.IsCompleted)
                    {
                        Console.WriteLine($"{watch.ElapsedMilliseconds,6} ms  still sending...");
                        await Task.WhenAny(sending, Task.Delay(200));
                    }

                    await sending;
                    Console.WriteLine($"{watch.ElapsedMilliseconds,6} ms  sent to {client.Username}");
                }
                catch (OperationCanceledException)
                {
                    Console.WriteLine("Gave up: sending took longer than 30 seconds.");
                }
            }
        }
    }
}
