// Demonstrates how to upload a message straight from an .eml file.
//
// AppendMessage has overloads that take the path of an .eml file, so a message saved
// earlier can be put on the server without loading it into a MailMessage yourself. The
// returned unique id identifies the new message; servers without UIDPLUS (RFC 4315)
// may not report one.

using System;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class AppendMessageFromFile
    {
        public static void Run()
        {
            var emlPath = Data.Imap/"MessageHeaders.eml";
            var folderName = "Aspose-" + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.CreateFolder(folderName);

                try
                {
                    var uid = client.AppendMessage(folderName, emlPath);
                    Console.WriteLine($"Uploaded {emlPath}");
                    Console.WriteLine($"  to '{folderName}', unique id: {uid ?? "(not reported)"}");

                    client.SelectFolder(folderName);
                    foreach (var info in client.ListMessages())
                    {
                        Console.WriteLine($"  subject: {info.Subject}");
                        Console.WriteLine($"  from:    {info.From}");
                        Console.WriteLine($"  size:    {info.Size} bytes");
                    }
                }
                finally
                {
                    client.DeleteFolder(folderName);
                }
            }
        }
    }
}
