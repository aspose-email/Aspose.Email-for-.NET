// Demonstrates how to find the folder a message belongs to, starting from the
// message's entry id.

using System;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.MAPI
{
    internal static class RetreiveParentFolderInformationFromMessageInfo
    {
        public static void Run()
        {
            // The source is only read, so open it read-only.
            using (var personalStorage = PersonalStorage.FromFile(Data.Mapi/"Outlook.pst", false))
            {
                var messages = 0;

                foreach (var folder in personalStorage.RootFolder.GetSubFolders())
                {
                    foreach (var msg in folder.EnumerateMessages())
                    {
                        // Useful when a message is reached through a search, where the
                        // folder it came from is not known.
                        var parent = personalStorage.GetParentFolder(msg.EntryId);

                        Console.WriteLine($"{msg.Subject}  ->  {parent.DisplayName}");
                        messages++;
                    }
                }

                Console.WriteLine($"\nResolved the parent folder of {messages} message(s).");
            }
        }
    }
}
