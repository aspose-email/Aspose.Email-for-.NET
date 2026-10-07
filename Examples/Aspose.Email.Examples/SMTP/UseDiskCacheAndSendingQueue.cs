// Demonstrates sending a batch through the client's disk-backed queue.
//
// SendToQueue stores the messages in SmtpQueueLocation (an absolute path) and returns,
// leaving the client to deliver them from there; SucceededSending and FailedSending
// report each outcome, possibly from other threads. Keeping the queue on disk means a
// large batch does not have to sit in memory.
//
// The example waits up to a minute for all outcomes and then shows what is left in the
// queue folder.

using System;
using System.IO;
using System.Linq;
using System.Threading;

namespace Aspose.Email.Examples.SMTP
{
    internal static class UseDiskCacheAndSendingQueue
    {
        public static void Run()
        {
            if (!ClientBuilder.IsSmtpConfigured)
            {
                SmtpExampleInfo.PrintNotConfigured();
                return;
            }

            const int messageCount = 5;
            var queueDir = Data.OutSub("SmtpQueue");

            var succeeded = 0;
            var failed = 0;

            using (var allDone = new ManualResetEventSlim(false))
            using (var client = ClientBuilder.Smtp(AuthType.Basic))
            {
                client.SmtpQueueLocation = queueDir;

                client.SucceededSending += (sender, e) =>
                {
                    if (Interlocked.Increment(ref succeeded) + Volatile.Read(ref failed) == messageCount)
                        allDone.Set();
                };

                client.FailedSending += (sender, e) =>
                {
                    Console.WriteLine($"  failed: {e.Message.Subject} - {e.OperationError?.Message}");
                    if (Interlocked.Increment(ref failed) + Volatile.Read(ref succeeded) == messageCount)
                        allDone.Set();
                };

                var messages = Enumerable.Range(1, messageCount)
                    .Select(i => new MailMessage(client.Username, client.Username, $"Queued message {i}", "Body"))
                    .ToList();

                client.SendToQueue(messages);
                Console.WriteLine($"Queued {messageCount} message(s) in {queueDir}");

                var finished = allDone.Wait(TimeSpan.FromMinutes(1));
                Console.WriteLine(finished
                    ? "All messages processed."
                    : "Stopped waiting after a minute.");
            }

            Console.WriteLine($"\n{succeeded} sent, {failed} failed.");
            var leftInQueue = Directory.GetFiles(queueDir, "*", SearchOption.AllDirectories).Length;
            Console.WriteLine($"Files left in the queue folder: {leftInQueue}");
        }
    }
}
