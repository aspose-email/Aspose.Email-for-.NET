// Demonstrates how to search an mbox storage with a query instead of reading every
// message and filtering afterwards.

using System;
using Aspose.Email.Storage.Mbox;
using Aspose.Email.Tools.Search;

namespace Aspose.Email.Examples.MBOX
{
    internal static class FilterMessagesInMbox
    {
        public static void Run()
        {
            using (var reader = MboxStorageReader.CreateReader(Data.Mbox/"ExampleMbox.mbox", new MboxLoadOptions()))
            {
                Console.WriteLine($"Messages in the storage: {reader.GetTotalItemsCount()}");

                // The mbox search engine works on the message headers, so SentDate is
                // supported here while InternalDate - a mailbox-side field - is not.
                var builder = new MailQueryBuilder();
                builder.Subject.Contains("a");
                builder.SentDate.Before(DateTime.Today);

                var matches = 0;
                foreach (var message in reader.EnumerateMessages(builder.GetQuery()))
                {
                    Console.WriteLine($"  {message.Subject}");
                    matches++;
                }

                Console.WriteLine($"\n{matches} message(s) matched the query.");
            }
        }
    }
}
