// Demonstrates how to open an Outlook PST file and read its basic properties.

using System;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.PST
{
    internal static class LoadingPSTFile
    {
        public static void Run()
        {
            // The source is only read, so open it read-only.
            using (var personalStorage = PersonalStorage.FromFile(Data.Mapi/"PersonalStorage.pst", false))
            {
                Console.WriteLine($"File format:    {personalStorage.Format}");
                Console.WriteLine($"Root folder:    {personalStorage.RootFolder.DisplayName}");
                Console.WriteLine($"Subfolders:     {personalStorage.RootFolder.GetSubFolders().Count}");
            }
        }
    }
}
