// Demonstrates downloading messages into memory instead of files.
//
// The SaveMessage overload that takes a Stream writes the MIME content there - handy for
// storing messages in a database, uploading them elsewhere or computing a hash, without
// touching the disk. MailMessage.Load parses the stream when needed.

using System;
using System.IO;

namespace Aspose.Email.Examples.EWS
{
    internal static class SaveMessagesToMemoryStreamUsingEWS
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                foreach (var info in client.ListMessages(client.MailboxInfo.InboxUri, 5))
                {
                    using (var stream = new MemoryStream())
                    {
                        client.SaveMessage(info.UniqueUri, stream);

                        stream.Position = 0;
                        var message = MailMessage.Load(stream);
                        Console.WriteLine($"{stream.Length,10:N0} bytes  {message.Subject}");
                    }
                }
            }
        }
    }
}
