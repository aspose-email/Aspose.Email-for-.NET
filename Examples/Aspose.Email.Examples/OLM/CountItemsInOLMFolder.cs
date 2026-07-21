// Demonstrates how to count the messages in every folder of an Outlook for Mac
// (OLM) storage file.

using System;
using System.Collections.Generic;
using Aspose.Email.Storage.Olm;

namespace Aspose.Email.Examples.OLM
{
    internal static class CountItemsInOLMFolder
    {
        public static void Run()
        {
            using (var storage = new OlmStorage(Data.Mapi/"SampleOLM.olm"))
            {
                PrintMessageCount(storage.FolderHierarchy);
            }
        }

        // FolderHierarchy lists only the top-level folders, so recurse to reach the rest.
        private static void PrintMessageCount(List<OlmFolder> folders, string indent = "")
        {
            foreach (var folder in folders)
            {
                Console.WriteLine($"{indent}{folder.Name}: {folder.MessageCount} message(s)");

                if (folder.SubFolders.Count > 0)
                    PrintMessageCount(folder.SubFolders, indent + "    ");
            }
        }
    }
}
