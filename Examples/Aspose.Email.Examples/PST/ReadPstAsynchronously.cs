// Demonstrates the asynchronous PST API: FromFileAsync and CreateAsync open or make a
// storage, SplitIntoAsync and MergeWithAsync do the long-running work. All of them take
// a CancellationToken so a job over a large storage can be abandoned.

using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.PST
{
    internal static class ReadPstAsynchronously
    {
        // The example runner calls a synchronous Run(), so bridge to the async body here.
        public static void Run()
        {
            RunAsync().GetAwaiter().GetResult();
        }

        private static async Task RunAsync()
        {
            var partsDir = Data.OutSub("PstAsyncParts");

            using (var cancellation = new CancellationTokenSource())
            {
                // Split an existing storage into chunks.
                using (var source = await PersonalStorage.FromFileAsync(Data.Mapi/"Sub.pst", false, cancellation.Token))
                {
                    Console.WriteLine($"Source items: {source.Store.GetTotalItemsCount()}");

                    await source.SplitIntoAsync(2000000, "Backup_", partsDir, cancellation.Token);
                }

                var parts = Directory.GetFiles(partsDir, "*.pst");
                Console.WriteLine($"Split into {parts.Length} part(s):");
                foreach (var part in parts)
                    Console.WriteLine($"  {Path.GetFileName(part)}  {new FileInfo(part).Length:N0} bytes");

                // Then build a new storage and merge the chunks back into it.
                var mergedPath = Data.Out/"ReadPstAsynchronously_merged.pst";

                using (var merged = await PersonalStorage.CreateAsync(
                           mergedPath, FileFormatVersion.Unicode, cancellation.Token))
                {
                    merged.CreatePredefinedFolder("Inbox", StandardIpmFolder.Inbox);

                    await merged.MergeWithAsync(parts, cancellation.Token);

                    Console.WriteLine($"\nMerged items: {merged.Store.GetTotalItemsCount()}");
                }

                Console.WriteLine($"Merged storage: {mergedPath}");
            }
        }
    }
}
