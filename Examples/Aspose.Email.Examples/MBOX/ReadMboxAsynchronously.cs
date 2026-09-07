// Demonstrates the asynchronous mbox API: CreateReaderAsync opens the storage and
// SplitIntoAsync breaks it up, both taking a CancellationToken so a long-running job
// over a large file can be abandoned.

using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Email.Storage.Mbox;

namespace Aspose.Email.Examples.MBOX
{
    internal static class ReadMboxAsynchronously
    {
        // The example runner calls a synchronous Run(), so bridge to the async body here.
        public static void Run() => RunAsync().GetAwaiter().GetResult();

        private static async Task RunAsync()
        {
            var mboxPath = BuildStorage(12);
            var outputDir = Data.OutSub("MboxAsyncParts");

            using (var cancellation = new CancellationTokenSource())
            using (var reader = await MboxStorageReader.CreateReaderAsync(
                       mboxPath, new MboxLoadOptions(), cancellation.Token))
            {
                Console.WriteLine($"Messages in the storage: {reader.GetTotalItemsCount()}");

                var parts = 0;
                reader.MboxFileCreated += (sender, e) => parts++;

                await reader.SplitIntoAsync(2000, outputDir, "Backup_", cancellation.Token);

                Console.WriteLine($"Split into {parts} part(s).");
            }

            foreach (var file in Directory.GetFiles(outputDir))
                Console.WriteLine($"  {Path.GetFileName(file)}  {new FileInfo(file).Length:N0} bytes");

            Console.WriteLine($"\nParts written to {outputDir}");
        }

        private static string BuildStorage(int messageCount)
        {
            var mboxPath = Data.Out/"ReadMboxAsynchronously_in.mbox";

            using (var writer = new MboxrdStorageWriter(mboxPath, new MboxSaveOptions()))
            {
                for (var i = 1; i <= messageCount; i++)
                {
                    writer.WriteMessage(new MailMessage(
                        "sender@domain.com", "receiver@domain.com", $"Message {i}", $"Body of message {i}"));
                }
            }

            return mboxPath;
        }
    }
}
