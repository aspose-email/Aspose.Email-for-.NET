// Demonstrates how to open an Outlook for Mac (OLM) storage file and read the
// messages out of its folders.

using System;
using Aspose.Email.Storage.Olm;

namespace Aspose.Email.Examples.OLM
{
    internal static class LoadAndReadOLMFile
    {
        public static void Run()
        {
            using (var storage = new OlmStorage(Data.Mapi/"SampleOLM.olm"))
            {
                var messages = 0;

                foreach (var folder in storage.FolderHierarchy)
                {
                    Console.WriteLine($"[{folder.Name}]");

                    // Messages are streamed out of the storage one folder at a time.
                    if (folder.HasMessages)
                    {
                        foreach (var msg in storage.EnumerateMessages(folder))
                        {
                            Console.WriteLine($"    Subject: {msg.Subject}");
                            messages++;
                        }
                    }

                    foreach (var subFolder in folder.SubFolders)
                        Console.WriteLine($"    Subfolder: {subFolder.Name}");
                }

                Console.WriteLine($"\nRead {messages} message(s).");
            }
        }
    }
}
