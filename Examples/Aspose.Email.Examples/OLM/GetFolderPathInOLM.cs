// Demonstrates how to read the full path of every folder in an Outlook for Mac
// (OLM) storage file.

using System;
using System.Collections.Generic;
using Aspose.Email.Storage.Olm;

namespace Aspose.Email.Examples.OLM
{
    internal static class GetFolderPathInOLM
    {
        public static void Run()
        {
            using (var storage = new OlmStorage(Data.Mapi/"SampleOLM.olm"))
            {
                var count = PrintPath(storage.FolderHierarchy);
                Console.WriteLine($"\nListed {count} folder(s).");
            }
        }

        // Each folder knows its own location, so there is no need to build the path up
        // while walking the hierarchy. Returns how many folders were printed.
        private static int PrintPath(List<OlmFolder> folders)
        {
            var count = 0;

            foreach (var folder in folders)
            {
                Console.WriteLine(folder.Path);
                count++;

                if (folder.SubFolders.Count > 0)
                    count += PrintPath(folder.SubFolders);
            }

            return count;
        }
    }
}
