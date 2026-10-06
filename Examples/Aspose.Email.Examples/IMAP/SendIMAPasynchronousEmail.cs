// Demonstrates downloading from several folders at the same time with the task-based
// API. Despite its historical name, this example does not send mail.
//
// One IMAP connection runs one command at a time, so concurrent downloads need separate
// connections. CreateConnection opens an extra one; passing it to the async methods - as
// the connection argument, or with SetConnection on a parameter set - makes those calls
// run over it, with a selected folder of their own. Dispose each connection when done.

using System;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Clients.Imap.Models;

namespace Aspose.Email.Examples.IMAP
{
    internal static class SendIMAPasynchronousEmail
    {
        public static void Run()
        {
            RunAsync().GetAwaiter().GetResult();
        }

        private static async Task RunAsync()
        {
            const int messagesPerFolder = 3;

            using (var imapClient = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                IAsyncImapClient client = imapClient;

                var folderNames = (await client.ListFoldersAsync())
                    .Where(folder => folder.Selectable)
                    .Select(folder => folder.Name)
                    .Take(3)
                    .ToList();

                // One download per folder, all running at once.
                var downloads = folderNames
                    .Select(folderName => Task.Run(() => DownloadNewestAsync(imapClient, folderName, messagesPerFolder)))
                    .ToArray();

                foreach (var report in await Task.WhenAll(downloads))
                    Console.WriteLine(report);
            }
        }

        private static async Task<string> DownloadNewestAsync(ImapClient imapClient, string folderName, int count)
        {
            IAsyncImapClient client = imapClient;

            using (var connection = imapClient.CreateConnection())
            {
                await client.SelectFolderAsync(folderName, connection: connection);
                var infos = await client.ListMessagesAsync(folderName, connection: connection);

                var newest = infos.OrderByDescending(info => info.InternalDate).Take(count).ToList();
                var report = new StringBuilder($"{folderName} (connection #{connection.ConnectionId}):");

                if (newest.Count == 0)
                    return report.Append(" empty").ToString();

                var messages = await client.FetchMessagesAsync(ImapFetchMessages.Create()
                    .SetMessages(newest)
                    .SetConnection(connection));

                foreach (var message in messages)
                    report.AppendLine().Append($"  {message.Date:g}  {message.Subject}");

                return report.ToString();
            }
        }
    }
}
