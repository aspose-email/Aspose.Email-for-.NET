// Demonstrates how to convert an Outlook offline storage (OST) file into a
// personal storage (PST) file.

using System;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.PST
{
    internal static class ConvertingOSTToPST
    {
        public static void Run()
        {
            var outputPath = Data.Out/"test.pst";

            // The source is only read, so open it read-only.
            using (var personalStorage = PersonalStorage.FromFile(Data.Mapi/"PersonalStorageFile.ost", false))
            {
                Console.WriteLine($"Source format: {personalStorage.Format}");

                personalStorage.SaveAs(outputPath, FileFormat.Pst);
            }

            Console.WriteLine($"Converted to PST: {outputPath}");
        }
    }
}
