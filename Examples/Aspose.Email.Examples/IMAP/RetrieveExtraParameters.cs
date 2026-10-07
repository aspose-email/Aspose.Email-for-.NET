// Demonstrates fetching additional, server-specific message attributes with the summary.
//
// ListMessage and ListMessages accept a list of extra FETCH items. Whatever the server
// returns for them ends up in ImapMessageInfo.ExtraParameters, keyed by the item name.
// Here they are Gmail's X-GM-MSGID (a message id stable across all folders) and
// X-GM-THRID (the conversation id), so the example needs a Gmail account.

using System;
using System.Linq;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class RetrieveExtraParameters
    {
        public static void Run()
        {
            var extraFields = new[] { "X-GM-MSGID", "X-GM-THRID" };

            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.SelectFolder(ImapFolderInfo.InBox);

                if (!client.GmExt1Supported)
                {
                    Console.WriteLine("The server does not support X-GM-EXT-1 (this example needs Gmail).");
                    return;
                }

                // A whole folder...
                var messages = client.ListMessages(extraFields);
                Console.WriteLine($"{messages.Count} message(s) listed with the extra fields, the first 5:");
                foreach (var info in messages.Take(5))
                    Print(info, extraFields);

                if (messages.Count == 0)
                    return;

                // ...or a single message, here by unique id; another overload takes a
                // sequence number.
                Console.WriteLine("\nOne message by unique id:");
                Print(client.ListMessage(messages[0].UniqueId, extraFields), extraFields);
            }
        }

        private static void Print(ImapMessageInfo info, string[] extraFields)
        {
            Console.WriteLine($"  {info.Subject}");
            foreach (var field in extraFields)
            {
                string value;
                info.ExtraParameters.TryGetValue(field, out value);
                Console.WriteLine($"    {field}: {value ?? "(not returned)"}");
            }
        }
    }
}
