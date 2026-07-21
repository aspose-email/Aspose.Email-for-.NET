// Demonstrates how to retrieve a specific range of messages from a PST folder by index and count.

using System;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.MAPI
{
    internal static class ExtractNumberOfMessages
    {
        public static void Run()
        {
            using (var pst = PersonalStorage.FromFile(Data.Mapi/"Sub.pst"))
            {
                var inbox = pst.RootFolder.GetSubFolder("Inbox");
                // Retrieve 100 messages starting from index 10
                var messages = inbox.GetContents(10, 100);
                Console.WriteLine($"Retrieved {messages.Count} messages.");
            }
        }
    }
}
