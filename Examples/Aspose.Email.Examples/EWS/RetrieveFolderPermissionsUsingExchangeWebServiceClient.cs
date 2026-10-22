// Demonstrates reading who has access to a folder, and how much.
//
// GetFolderPermissions returns one entry per user (plus the Default and Anonymous
// entries) with the permission level and the individual rights - read, create, edit,
// delete, folder owner. This is the "Permissions" tab of a folder in Outlook. The
// example reads the permissions of the Calendar and Contacts folders.

using System;
using Aspose.Email.Clients.Exchange;
using Aspose.Email.Clients.Exchange.WebService;

namespace Aspose.Email.Examples.EWS
{
    internal static class RetrieveFolderPermissionsUsingExchangeWebServiceClient
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                Print(client, "Calendar", client.MailboxInfo.CalendarUri);
                Print(client, "Contacts", client.MailboxInfo.ContactsUri);
            }
        }

        private static void Print(IEWSClient client, string name, string folderUri)
        {
            Console.WriteLine($"{name}:");

            foreach (var entry in client.GetFolderPermissions(folderUri))
            {
                var user = entry.UserInfo?.PrimarySmtpAddress ?? entry.UserInfo?.DisplayName ?? "(default)";
                var permission = entry as ExchangeFolderPermission;

                if (permission == null)
                {
                    Console.WriteLine($"  {user}: {entry.GetType().Name}");
                    continue;
                }

                Console.WriteLine($"  {user}: {permission.PermissionLevel}");
                Console.WriteLine($"    read: {permission.ReadItems}, create: {permission.CanCreateItems}, " +
                                  $"delete: {permission.DeleteItems}, owner: {permission.IsFolderOwner}, " +
                                  $"visible: {permission.IsFolderVisible}");
            }
        }
    }
}
