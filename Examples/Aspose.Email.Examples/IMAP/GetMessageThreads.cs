// Demonstrates server-side threading with the THREAD extension (RFC 5256).
//
// The server groups the messages of the selected folder into conversations and returns
// each one as a tree: a MessageThreadResult is a message, its ChildMessages are the
// replies to it. REFERENCES threads by the In-Reply-To/References headers,
// ORDEREDSUBJECT by subject; ThreadAlgorithms lists what the server offers.

using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class GetMessageThreads
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.SelectFolder(ImapFolderInfo.InBox);

                if (!client.ThreadSupported || client.ThreadAlgorithms.Length == 0)
                {
                    Console.WriteLine("The server does not support THREAD.");
                    return;
                }

                Console.WriteLine($"Thread algorithms: {string.Join(", ", client.ThreadAlgorithms)}");

                var algorithm = client.ThreadAlgorithms.Contains("REFERENCES") ? "REFERENCES" : client.ThreadAlgorithms[0];

                var conditions = new ThreadSearchConditions
                {
                    Algorithm = algorithm,
                    UseUId = true,
                    Since = DateTime.Today.AddDays(-30)
                };

                var threads = client.GetMessageThreads(conditions);
                Console.WriteLine($"{threads.Count} thread(s) in the last 30 days, grouped by {algorithm}.\n");

                // Look up the subjects of the first ten threads in one round trip.
                var shown = threads.Take(10).ToList();
                var uids = new List<string>();
                foreach (var thread in shown)
                    CollectUids(thread, uids);

                var subjects = client.ListMessages(ImapFolderInfo.InBox, uids)
                    .ToDictionary(info => info.UniqueId, info => info.Subject);

                foreach (var thread in shown)
                    Print(thread, subjects, 0);
            }
        }

        private static void CollectUids(MessageThreadResult node, List<string> uids)
        {
            uids.Add(node.UniqueId);
            foreach (var child in node.ChildMessages)
                CollectUids(child, uids);
        }

        private static void Print(MessageThreadResult node, IDictionary<string, string> subjects, int depth)
        {
            string subject;
            subjects.TryGetValue(node.UniqueId, out subject);

            Console.WriteLine($"{new string(' ', 2 + depth * 2)}uid {node.UniqueId}: {subject}");
            foreach (var child in node.ChildMessages)
                Print(child, subjects, depth + 1);
        }
    }
}
