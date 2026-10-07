// Demonstrates custom flags (IMAP keywords), which tag messages with your own labels.
//
// ImapMessageFlags.Keyword creates a flag with any name; it is stored on the server like
// the standard flags, so every client sees it. ContainsKeyword checks for it, and
// HasFlags searches for it. A server lists whether it accepts new keywords in its
// PERMANENTFLAGS; a few do not.
// The example works in a uniquely named folder and deletes it at the end.

using System;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class SetCustomFlag
    {
        public static void Run()
        {
            var folderName = "Aspose-" + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.CreateFolder(folderName);

                try
                {
                    var tagged = client.AppendMessage(folderName,
                        new MailMessage("sender@example.com", "receiver@example.com", "Contract draft", "Body"));
                    client.AppendMessage(folderName,
                        new MailMessage("sender@example.com", "receiver@example.com", "Lunch?", "Body"));

                    client.SelectFolder(folderName);
                    var legal = ImapMessageFlags.Keyword("legal");
                    client.AddMessageFlags(tagged, legal | ImapMessageFlags.Keyword("urgent"));

                    foreach (var info in client.ListMessages())
                    {
                        Console.WriteLine($"{info.Subject,-16} flags: {info.Flags}, " +
                                          $"has 'legal': {info.ContainsKeyword("legal")}");
                    }

                    var builder = new ImapQueryBuilder();
                    builder.HasFlags(legal);

                    var found = client.ListMessages(builder.GetQuery());
                    Console.WriteLine($"\nSearch for keyword 'legal': {found.Count} message(s)");
                }
                finally
                {
                    client.DeleteFolder(folderName);
                }
            }
        }
    }
}
