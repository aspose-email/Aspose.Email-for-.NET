// Demonstrates how to use one ImapClient from several threads at once.
//
// An IMAP connection runs one command at a time and has a single selected folder, so
// threads sharing it would trip over each other. CreateConnection opens an extra,
// independent connection: pass it to any overload that takes an IConnection and the
// call runs over that connection, with its own selected folder. Dispose it when done.

using System;
using System.Linq;
using System.Threading.Tasks;

namespace Aspose.Email.Examples.IMAP
{
    internal static class WorkWithMultipleConnections
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                var folderNames = client.ListFolders()
                    .Where(folder => folder.Selectable)
                    .Select(folder => folder.Name)
                    .Take(3)
                    .ToList();

                // One task per folder, each on a connection of its own.
                var tasks = folderNames.Select(folderName => Task.Run(() =>
                {
                    using (var connection = client.CreateConnection())
                    {
                        client.SelectFolder(connection, folderName);
                        var messages = client.ListMessages(connection);
                        return $"{folderName}: {messages.Count} message(s), connection #{connection.ConnectionId}";
                    }
                })).ToArray();

                Task.WaitAll(tasks);

                foreach (var task in tasks)
                    Console.WriteLine(task.Result);
            }
        }
    }
}
