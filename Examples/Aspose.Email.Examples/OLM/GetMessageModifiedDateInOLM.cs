// Demonstrates how to read the summary properties of OLM messages, including the
// modification date, without extracting the messages themselves.

using System;
using System.Collections.Generic;
using Aspose.Email.Storage.Olm;

namespace Aspose.Email.Examples.OLM
{
    internal static class GetMessageModifiedDateInOLM
    {
        public static void Run()
        {
            using (var olm = new OlmStorage(Data.Mapi/"SampleOLM.olm"))
            {
                PrintMessages(olm.FolderHierarchy);
            }
        }

        private static void PrintMessages(List<OlmFolder> folders)
        {
            foreach (var folder in folders)
            {
                if (folder.HasMessages)
                {
                    Console.WriteLine($"[{folder.Name}]");

                    // EnumerateMessages on the folder yields OlmMessageInfo, which carries
                    // the summary properties only - no message is parsed here.
                    foreach (var messageInfo in folder.EnumerateMessages())
                    {
                        Console.WriteLine($"  {messageInfo.Subject}");
                        Console.WriteLine($"    sent:     {messageInfo.Date}");
                        Console.WriteLine($"    modified: {messageInfo.ModifiedDate}");
                    }
                }

                if (folder.SubFolders.Count > 0)
                    PrintMessages(folder.SubFolders);
            }
        }
    }
}
