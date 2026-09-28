// Demonstrates Gmail's own conversation threading (X-GM-EXT-1).
//
// Gmail does not implement the THREAD extension. Instead every message carries the id
// of the conversation it belongs to (X-GM-THRID), exposed as ImapMessageInfo.
// ConversationId. XGMThreadSearchConditions finds all messages of one conversation in
// the selected folder - use "[Gmail]/All Mail" to include your own replies as well.

using System;
using System.Linq;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class GetGmailConversationThreads
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.SelectFolder(ImapFolderInfo.InBox);

                if (!client.GmExt1Supported)
                {
                    Console.WriteLine("The server does not support X-GM-EXT-1 (this example needs Gmail).");
                    return;
                }

                // Ask for the conversation id along with the usual summary fields.
                var messages = client.ListMessages(ImapFolderInfo.InBox, false, new[] { "X-GM-THRID" });

                var latest = messages
                    .Where(info => !string.IsNullOrEmpty(GetConversationId(info)))
                    .OrderByDescending(info => info.InternalDate)
                    .FirstOrDefault();

                if (latest == null)
                {
                    Console.WriteLine("No message with a conversation id found.");
                    return;
                }

                var conversationId = GetConversationId(latest);
                Console.WriteLine($"Conversation {conversationId}, started from: {latest.Subject}");

                var conditions = new XGMThreadSearchConditions
                {
                    ConversationId = conversationId,
                    UseUId = true
                };

                var thread = client.GetMessageThreads(conditions);
                var uids = thread.Select(result => result.UniqueId).ToList();

                Console.WriteLine($"{uids.Count} message(s) in this folder belong to it:");
                foreach (var info in client.ListMessages(ImapFolderInfo.InBox, uids))
                    Console.WriteLine($"  {info.InternalDate:g}  {info.From}  {info.Subject}");
            }
        }

        private static string GetConversationId(ImapMessageInfo info)
        {
            if (!string.IsNullOrEmpty(info.ConversationId))
                return info.ConversationId;

            string value;
            return info.ExtraParameters != null && info.ExtraParameters.TryGetValue("X-GM-THRID", out value)
                ? value
                : null;
        }
    }
}
