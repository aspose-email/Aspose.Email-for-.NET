// Demonstrates filtering and paging the messages of an OLM folder.
//
// A folder reads from the storage as a forward-only cursor. Draining an enumeration to
// the end rewinds it, but stopping early (Take, First, a break) leaves it part-way
// through, so the next enumeration carries on from there instead of starting over.
// Anything that reads a slice therefore opens its own storage.

using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Email.Storage.Olm;
using Aspose.Email.Tools.Search;

namespace Aspose.Email.Examples.OLM
{
    internal static class FilterMessagesInOLMFolder
    {
        private const string OlmFile = "SampleOLM.olm";

        public static void Run()
        {
            string folderPath;
            int total;

            using (var storage = new OlmStorage(Data.Mapi/OlmFile))
            {
                var folder = FindLargestFolder(storage.GetFolders());
                if (folder == null)
                {
                    Console.WriteLine("No folder in this storage holds any messages.");
                    return;
                }

                folderPath = folder.Path;
                total = folder.MessageCount;
                Console.WriteLine($"Folder: {folderPath} ({total} message(s))");

                var builder = new MailQueryBuilder();
                builder.Subject.Contains("message");

                Console.WriteLine("\nMatching the query:");
                foreach (var info in folder.EnumerateMessages(builder.GetQuery()))
                    Console.WriteLine($"  {info.Subject}");
            }

            const int pageSize = 2;

            for (var startIndex = 0; startIndex < total; startIndex += pageSize)
            {
                // A fresh storage per page: reusing one would resume mid-folder.
                using (var storage = new OlmStorage(Data.Mapi/OlmFile))
                {
                    var folder = FindByPath(storage.GetFolders(), folderPath);

                    Console.WriteLine($"\n--- messages {startIndex}..{Math.Min(startIndex + pageSize, total) - 1} ---");
                    foreach (var info in folder.EnumerateMessages(startIndex, pageSize))
                        Console.WriteLine($"  {info.Subject}");
                }
            }
        }

        private static OlmFolder FindLargestFolder(List<OlmFolder> folders)
        {
            OlmFolder best = null;

            foreach (var folder in folders)
            {
                if (best == null || folder.MessageCount > best.MessageCount)
                    best = folder;

                if (folder.SubFolders != null && folder.SubFolders.Count > 0)
                {
                    var deeper = FindLargestFolder(folder.SubFolders);
                    if (deeper != null && (best == null || deeper.MessageCount > best.MessageCount))
                        best = deeper;
                }
            }

            return best != null && best.MessageCount > 0 ? best : null;
        }

        private static OlmFolder FindByPath(List<OlmFolder> folders, string path)
        {
            foreach (var folder in folders)
            {
                if (folder.Path == path)
                    return folder;

                if (folder.SubFolders != null && folder.SubFolders.Count > 0)
                {
                    var deeper = FindByPath(folder.SubFolders, path);
                    if (deeper != null)
                        return deeper;
                }
            }

            return null;
        }
    }
}
