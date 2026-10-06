// Demonstrates how to connect to Microsoft Graph and read the mailbox folder tree.
//
// Graph is OAuth-only: there is no password to pass. The client is built from an
// application registration in Entra ID, and with application permissions the token
// belongs to the app rather than to a person - so the mailbox has to be named through
// Resource and ResourceId. Put your registration into clientsettings.json.

using System;
using Aspose.Email.Clients.Graph;

namespace Aspose.Email.Examples.Graph
{
    internal static class ConnectToGraphClient
    {
        public static void Run()
        {
            if (!ClientBuilder.IsGraphConfigured)
            {
                GraphExampleInfo.PrintNotConfigured();
                return;
            }

            using (var client = ClientBuilder.Graph(AuthType.ModernWithAppPermission))
            {
                // The endpoint only needs setting for a sovereign cloud - for example
                // "https://graph.microsoft.us" for GCC High, or the default commercial
                // "https://graph.microsoft.com".
                Console.WriteLine($"Endpoint:    {client.EndPoint}");
                Console.WriteLine($"Tenant:      {client.TenantId}");
                Console.WriteLine($"Resource:    {client.Resource} / {client.ResourceId}");

                // Timeout is in milliseconds and applies to each request.
                client.Timeout = 60000;

                // A proxy can be supplied when the network requires one:
                //   client.Proxy = new WebProxy("http://proxy.example.com:8080");

                var folders = client.ListFolders(null);
                Console.WriteLine($"\n{folders.Count} folder(s) in the mailbox:");

                foreach (var folder in folders)
                {
                    Console.WriteLine($"  {folder.DisplayName}: {folder.ContentCount} item(s), " +
                                      $"{folder.ContentUnreadCount} unread, subfolders: {folder.HasSubFolders}");
                }

                // Well-known folders are addressed by name instead of by id.
                var inbox = client.GetFolder(KnownFolders.Inbox);
                Console.WriteLine($"\nInbox id: {inbox.ItemId}");
            }
        }
    }
}
