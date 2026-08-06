// Demonstrates how to list the summary properties of every message in an mbox
// storage. EnumerateMessageInfo reads the headers only, so nothing is parsed twice.

using System;
using Aspose.Email.Storage.Mbox;

namespace Aspose.Email.Examples.MBOX
{
    internal static class ReadMessageInfoFromMbox
    {
        public static void Run()
        {
            using (var reader = MboxStorageReader.CreateReader(Data.Mbox/"ExampleMbox.mbox", new MboxLoadOptions()))
            {
                foreach (var messageInfo in reader.EnumerateMessageInfo())
                {
                    Console.WriteLine($"Subject: {messageInfo.Subject}");
                    Console.WriteLine($"Date:    {messageInfo.Date}");
                    Console.WriteLine($"From:    {messageInfo.From}");
                    Console.WriteLine($"To:      {messageInfo.To}");
                    Console.WriteLine($"CC:      {messageInfo.CC}");
                    Console.WriteLine($"Bcc:     {messageInfo.Bcc}");
                    Console.WriteLine("-----------------------------------");
                }
            }
        }
    }
}
