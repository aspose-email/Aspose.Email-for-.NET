// Demonstrates how to size up a Zimbra TGZ backup before exporting it, so a progress
// indicator has something to count against.

using System;
using Aspose.Email.Storage.Zimbra;

namespace Aspose.Email.Examples.Email
{
    internal static class GetTotalItemsCountFromTgz
    {
        public static void Run()
        {
            using (var reader = new TgzReader(Data.Email/"ZimbraSample.tgz"))
            {
                var count = reader.GetTotalItemsCount();
                Console.WriteLine($"Items in the archive: {count}");
            }

            // A TGZ backup also carries contacts and calendar items; ExportTo writes them
            // out alongside the messages, keeping the archive's folder structure.
            var outputDir = Data.OutSub("ZimbraAllItems");

            using (var reader = new TgzReader(Data.Email/"ZimbraSample.tgz"))
            {
                reader.ExportTo(outputDir);
            }

            var files = System.IO.Directory.GetFiles(outputDir, "*", System.IO.SearchOption.AllDirectories);
            Console.WriteLine($"Exported {files.Length} file(s) to {outputDir}");
        }
    }
}
