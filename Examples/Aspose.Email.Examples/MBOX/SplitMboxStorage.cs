// Demonstrates how to break a large mbox storage into smaller parts and follow the
// progress through the reader's events.

using System;
using System.IO;
using Aspose.Email.Storage.Mbox;

namespace Aspose.Email.Examples.MBOX
{
    internal static class SplitMboxStorage
    {
        public static void Run()
        {
            var outputDir = Data.OutSub("MboxParts");

            using (var mbox = MboxStorageReader.CreateReader(Data.Mbox/"ExampleMbox.mbox",
                       new MboxLoadOptions { LeaveOpen = false }))
            {
                var messageCount = 0;
                var partCount = 0;

                mbox.MboxFileCreated += (sender, e) =>
                {
                    Console.WriteLine($"New mbox file created: {e.FileName}");
                    partCount++;
                };

                mbox.MboxFileFilled += (sender, e) =>
                {
                    Console.WriteLine($"Mbox file filled: {e.FileName}");
                };

                mbox.EmlCopied += (sender, e) =>
                {
                    Console.WriteLine($"  copied: {e.Item.Subject}");
                    messageCount++;
                };

                // A small chunk size forces several parts out of the sample storage.
                mbox.SplitInto(20000, outputDir, "Part");

                Console.WriteLine($"\n{messageCount} message(s) copied into {partCount} part(s).");
            }

            Console.WriteLine($"{Directory.GetFiles(outputDir).Length} file(s) written to {outputDir}");
        }
    }
}
