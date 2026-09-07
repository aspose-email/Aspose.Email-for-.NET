// Demonstrates the fault-tolerant way to walk an OLM storage.
//
// The parameterless-path constructors throw as soon as anything in the file cannot be
// parsed, which aborts the whole traversal. The callback constructor instead reports
// each failure - with the id of the item that caused it - and carries on, so one bad
// item does not cost you the rest of the storage.

using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Email.Storage.Olm;

namespace Aspose.Email.Examples.OLM
{
    internal static class ReadOLMWithErrorHandling
    {
        public static void Run()
        {
            var failures = new List<string>();

            // The callback receives the exception and the id of the offending item.
            using (var storage = new OlmStorage((exception, itemId) =>
                   {
                       failures.Add($"{itemId}: {exception.Message}");
                   }))
            {
                // Load returns false rather than throwing when the file cannot be opened.
                if (!storage.Load(Data.Mapi/"SampleOLM.olm"))
                {
                    Console.WriteLine("The storage could not be loaded.");
                    return;
                }

                Console.WriteLine($"Total items: {storage.GetTotalItemsCount()}");

                var messages = 0;
                foreach (var folder in Flatten(storage.GetFolders()))
                {
                    if (!folder.HasMessages)
                        continue;

                    Console.WriteLine($"[{folder.Path}]");

                    // EnumerateMapiMessages yields the messages themselves rather than
                    // their summaries, so nothing has to be extracted afterwards.
                    foreach (var message in folder.EnumerateMapiMessages())
                    {
                        Console.WriteLine($"  {message.Subject}");
                        messages++;
                    }
                }

                Console.WriteLine($"\nRead {messages} message(s).");
                Console.WriteLine($"Items skipped because of errors: {failures.Count}");

                foreach (var failure in failures.Take(10))
                    Console.WriteLine($"  {failure}");
            }
        }

        private static IEnumerable<OlmFolder> Flatten(List<OlmFolder> folders)
        {
            foreach (var folder in folders)
            {
                yield return folder;

                if (folder.SubFolders == null)
                    continue;

                foreach (var child in Flatten(folder.SubFolders))
                    yield return child;
            }
        }
    }
}
