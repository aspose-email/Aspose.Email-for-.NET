// Demonstrates how to follow a long-running PST split: the StorageProcessing event
// reports each part file as it is created.

using System;
using System.IO;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.MAPI
{
    internal static class TrackPSTStorageProcessing
    {
        public static void Run()
        {
            var outputDir = Data.OutSub("TrackPSTStorageProcessing");

            using (var pst = PersonalStorage.FromFile(Data.Mapi/"Sub.pst", false))
            {
                var partCount = 0;

                pst.StorageProcessing += (sender, args) =>
                {
                    Console.WriteLine($"Storage processing: {args.FileName}");
                    partCount++;
                };

                // A chunk size well below the source size forces several parts out of it.
                // It cannot go below the minimum size of a PST file.
                pst.SplitInto(2000000, "part_", outputDir);

                Console.WriteLine($"\n{partCount} event(s) raised.");
            }

            var files = Directory.GetFiles(outputDir, "*.pst");
            Console.WriteLine($"{files.Length} part file(s) written to {outputDir}");
        }
    }
}
