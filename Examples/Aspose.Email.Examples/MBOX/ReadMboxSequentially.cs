// Demonstrates the reader's cursor API. NextMessage walks the storage one header block
// at a time and hands back the entry id, which ExtractMessage then turns into a full
// message - useful when the decision to load a message depends on its headers.
//
// The cursor is forward-only: it never rewinds, so a second pass needs a second reader.

using System;
using Aspose.Email.Storage.Mbox;

namespace Aspose.Email.Examples.MBOX
{
    internal static class ReadMboxSequentially
    {
        public static void Run()
        {
            var mboxPath = BuildStorage();

            using (var reader = MboxStorageReader.CreateReader(mboxPath, new MboxLoadOptions()))
            {
                Console.WriteLine($"Storage size: {reader.BaseStream.Length:N0} bytes");
                Console.WriteLine();

                MboxMessageInfo info;
                var index = 0;

                while ((info = reader.NextMessage()) != null)
                {
                    Console.WriteLine($"{index}. {info.Subject}");
                    Console.WriteLine($"   from:      {info.From}");
                    Console.WriteLine($"   date:      {info.Date:yyyy-MM-dd HH:mm}");
                    Console.WriteLine($"   delimiter: {info.DelimiterMark.Trim()}");
                    Console.WriteLine($"   headers:   {info.Headers.Count}");

                    // Only the messages worth reading in full are extracted, and the
                    // extraction uses a separate reader because the cursor cannot rewind.
                    if (info.Subject.EndsWith("2", StringComparison.Ordinal))
                    {
                        using (var fetcher = MboxStorageReader.CreateReader(mboxPath, new MboxLoadOptions()))
                        {
                            var message = fetcher.ExtractMessage(info.EntryId, new EmlLoadOptions());
                            Console.WriteLine($"   body:      {message.Body.Trim()}");
                        }
                    }

                    index++;
                }

                // CurrentDataSize counts only what ReadNextMessage consumed, so the
                // cursor above leaves it at zero - no body was ever read.
                Console.WriteLine($"\nWalked {index} message(s) without loading them.");
            }
        }

        private static string BuildStorage()
        {
            var mboxPath = Data.Out/"ReadMboxSequentially_in.mbox";

            using (var writer = new MboxrdStorageWriter(mboxPath, new MboxSaveOptions()))
            {
                for (var i = 1; i <= 4; i++)
                {
                    writer.WriteMessage(new MailMessage(
                        $"sender{i}@domain.com", "receiver@domain.com", $"Message {i}", $"Body of message {i}")
                    {
                        Date = new DateTime(2024, 1, i, 10, 0, 0, DateTimeKind.Utc)
                    });
                }
            }

            return mboxPath;
        }
    }
}
