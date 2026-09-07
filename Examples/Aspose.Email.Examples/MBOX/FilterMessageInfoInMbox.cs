// Demonstrates filtering and paging over message *metadata*. EnumerateMessageInfo
// reads the headers only, so picking messages by subject or date never pays the cost
// of parsing bodies and attachments - only the messages that survive the filter are
// loaded in full.

using System;
using Aspose.Email.Storage.Mbox;
using Aspose.Email.Tools.Search;

namespace Aspose.Email.Examples.MBOX
{
    internal static class FilterMessageInfoInMbox
    {
        public static void Run()
        {
            var mboxPath = BuildStorage();

            var builder = new MailQueryBuilder();
            builder.Subject.Contains("Report");
            var query = builder.GetQuery();

            // Headers only - no body is parsed here.
            Console.WriteLine("Matching message info:");
            using (var reader = MboxStorageReader.CreateReader(mboxPath, new MboxLoadOptions()))
            {
                foreach (var info in reader.EnumerateMessageInfo(query))
                    Console.WriteLine($"  {info.Date:yyyy-MM-dd}  {info.From}  {info.Subject}");
            }

            // The same slicing as EnumerateMessages, still metadata only. A reader is a
            // forward-only cursor, so each slice needs its own reader.
            Console.WriteLine("\nSecond and third message info:");
            using (var reader = MboxStorageReader.CreateReader(mboxPath, new MboxLoadOptions()))
            {
                foreach (var info in reader.EnumerateMessageInfo(1, 2))
                    Console.WriteLine($"  {info.Subject}  (entry id {info.EntryId})");
            }

            // When the bodies are needed too, the query can be combined with load options
            // so only the matching messages are parsed.
            Console.WriteLine("\nMatching messages, loaded in full:");
            using (var reader = MboxStorageReader.CreateReader(mboxPath, new MboxLoadOptions()))
            {
                var loadOptions = new EmlLoadOptions { PreserveTnefAttachments = true };

                foreach (var message in reader.EnumerateMessages(loadOptions, query))
                    Console.WriteLine($"  {message.Subject}: {message.Body.Trim()}");
            }
        }

        private static string BuildStorage()
        {
            var mboxPath = Data.Out/"FilterMessageInfoInMbox_in.mbox";

            using (var writer = new MboxrdStorageWriter(mboxPath, new MboxSaveOptions()))
            {
                for (var i = 1; i <= 6; i++)
                {
                    var subject = (i % 2 == 0 ? "Report " : "Notice ") + i;

                    writer.WriteMessage(new MailMessage(
                        $"sender{i}@domain.com", "receiver@domain.com", subject, $"Body of message {i}")
                    {
                        Date = new DateTime(2024, 1, i, 10, 0, 0, DateTimeKind.Utc)
                    });
                }
            }

            return mboxPath;
        }
    }
}
