// Demonstrates the three ways to reach a folder in an OLM storage: the flat list from
// GetFolders, a lookup by name with GetFolder, and a lookup among a folder's own
// children with GetSubFolder.
//
// Worth knowing: OlmStorage.FromFile leaves FolderHierarchy null until the folders are
// actually read, whereas the constructor populates it. GetFolders works either way.

using System;
using System.Collections.Generic;
using Aspose.Email.Storage.Olm;

namespace Aspose.Email.Examples.OLM
{
    internal static class FindFolderInOLM
    {
        public static void Run()
        {
            using (var storage = OlmStorage.FromFile(Data.Mapi/"SampleOLM.olm"))
            {
                Console.WriteLine($"FolderHierarchy straight after FromFile: " +
                                  (storage.FolderHierarchy == null ? "not read yet" : "populated"));

                // GetFolders reads the hierarchy and returns the top-level folders.
                var folders = storage.GetFolders();
                Console.WriteLine($"GetFolders() returned {folders.Count} top-level folder(s)");
                Console.WriteLine($"FolderHierarchy now: " +
                                  (storage.FolderHierarchy == null ? "not read yet" : $"{storage.FolderHierarchy.Count} folder(s)"));

                Console.WriteLine("\nTop-level folders:");
                foreach (var folder in folders)
                    Console.WriteLine($"  {folder.Path} ({folder.MessageCount} message(s), {folder.SubFolders.Count} subfolder(s))");

                // A lookup by name, case-insensitively.
                var inbox = storage.GetFolder("Inbox", true);
                Console.WriteLine($"\nGetFolder(\"Inbox\", true) -> {(inbox == null ? "not found" : inbox.Path)}");

                // Nested folders are reached through their parent.
                var parent = FindWithSubfolders(folders);
                if (parent != null)
                {
                    var childName = parent.SubFolders[0].Name;
                    var child = parent.GetSubFolder(childName, true);

                    Console.WriteLine($"\n{parent.Name}.GetSubFolder(\"{childName}\", true) -> " +
                                      (child == null ? "not found" : child.Path));
                }
            }
        }

        private static OlmFolder FindWithSubfolders(List<OlmFolder> folders)
        {
            foreach (var folder in folders)
            {
                if (folder.SubFolders != null && folder.SubFolders.Count > 0)
                    return folder;
            }

            return null;
        }
    }
}
