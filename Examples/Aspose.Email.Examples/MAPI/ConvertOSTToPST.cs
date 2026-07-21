// Demonstrates how to convert an Outlook OST storage file to PST format.

using System;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.MAPI
{
    internal static class ConvertOstToPst
    {
        public static void Run()
        {
            var outputPath = Data.Out/"ConvertOSTToPST_out.pst";

            // The source is only read, so open it read-only.
            using (var ost = PersonalStorage.FromFile(Data.Mapi/"SampleOstFile.ost", false))
            {
                Console.WriteLine($"Source format: {ost.Format}");

                // SaveAs rewrites the whole storage in the target format.
                ost.SaveAs(outputPath, FileFormat.Pst);
            }

            Console.WriteLine($"Converted to PST: {outputPath}");
        }
    }
}
