// Demonstrates converting an OST into a PST in place.
//
// ConvertTo changes the format of the storage that is already open, so the file that
// was loaded becomes a PST. That differs from SaveAs, which leaves the source alone and
// writes a second file.

using System;
using System.IO;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.PST
{
    internal static class ConvertOstToPstInPlace
    {
        public static void Run()
        {
            // Work on a copy: the conversion rewrites the file it was opened from.
            var path = Data.Out/"ConvertOstToPstInPlace_out.pst";
            File.Copy(Data.Mapi/"SampleOstFile.ost", path, true);

            using (var storage = PersonalStorage.FromFile(path))
            {
                Console.WriteLine($"Format before: {storage.Format}");
                Console.WriteLine($"Items:         {storage.Store.GetTotalItemsCount()}");

                storage.ConvertTo(FileFormat.Pst);

                Console.WriteLine($"Format after:  {storage.Format}");
            }

            // Reopen to confirm the change was written to disk, not just to the instance.
            using (var storage = PersonalStorage.FromFile(path, false))
            {
                Console.WriteLine($"Reopened as:   {storage.Format}");
                Console.WriteLine($"Items:         {storage.Store.GetTotalItemsCount()}");
            }

            Console.WriteLine($"\nSaved to {path}");
        }
    }
}
