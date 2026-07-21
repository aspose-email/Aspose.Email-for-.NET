// Demonstrates how to create a MAPI journal entry and add it to a PST journal folder.

using System;
using System.IO;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.MAPI
{
    internal static class CreateNewMapiJournalAndAddToSubfolder
    {
        public static void Run()
        {
            var journal = new MapiJournal("daily record", "called out in the dark", "Phone call", "Phone call");
            journal.StartTime = DateTime.Now;
            journal.EndTime = journal.StartTime.AddHours(1);

            var path = Data.Out/"CreateNewMapiJournalAndAddToSubfolder_out.pst";

            if (File.Exists(path))
                File.Delete(path);

            using (var pst = PersonalStorage.Create(path, FileFormatVersion.Unicode))
            {
                var journalFolder = pst.CreatePredefinedFolder("Journal", StandardIpmFolder.Journal);
                journalFolder.AddMapiMessageItem(journal);

                Console.WriteLine($"Journal: {journal.Subject} ({journal.Description})");
                Console.WriteLine($"Logged:  {journal.StartTime:g} - {journal.EndTime:t} ({journal.Duration})");
                Console.WriteLine($"Saved to {path}");
            }
        }
    }
}
