// Demonstrates that OST files are opened with the same API as PST files -
// PersonalStorage detects the format itself.

using System;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.PST
{
    internal static class ReadingOSTFiles
    {
        public static void Run()
        {
            // The source is only read, so open it read-only.
            using (var personalStorage = PersonalStorage.FromFile(Data.Mapi/"PersonalStorageFile.ost", false))
            {
                Console.WriteLine($"File format of OST: {personalStorage.Format}");

                foreach (var folder in personalStorage.RootFolder.GetSubFolders())
                    Console.WriteLine($"    {folder.DisplayName}: {folder.ContentCount} item(s)");
            }
        }
    }
}
