// Demonstrates how to iterate all messages in a Zimbra TGZ archive
// and print each message's directory path and subject.

using System;
using Aspose.Email.Storage.Zimbra;

namespace Aspose.Email.Examples.Email
{
    internal static class ReadAllMessagesFromZimbraTgzStorage
    {
        public static void Run()
        {
            var reader = new TgzReader(Data.Email/"ZimbraSample.tgz");

            while (reader.ReadNextMessage())
            {
                Console.WriteLine(reader.CurrentDirectory);
                Console.WriteLine(reader.CurrentMessage.Subject);
            }
        }
    }
}
