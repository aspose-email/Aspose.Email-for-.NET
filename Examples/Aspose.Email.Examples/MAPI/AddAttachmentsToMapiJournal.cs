// Demonstrates how to create a MAPI journal entry with file attachments and save it as MSG.

using System;
using System.IO;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class AddAttachmentsToMapiJournal
    {
        public static void Run()
        {
            var journal = new MapiJournal("testJournal", "This is a test journal", "Phone call", "Phone call")
            {
                StartTime = DateTime.Now,
                Companies = new string[] { "company 1", "company 2", "company 3" }
            };
            journal.EndTime = journal.StartTime.AddHours(1);

            foreach (var file in new[] { Data.Mapi/"Desert.jpg", Data.Mapi/"download.png" })
                journal.Attachments.Add(file, File.ReadAllBytes(file));

            var outputPath = Data.Out/"AddAttachmentsToMapiJournal_out.msg";
            journal.Save(outputPath);

            Console.WriteLine($"Journal: {journal.Subject} ({journal.Attachments.Count} attachment(s))");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}
