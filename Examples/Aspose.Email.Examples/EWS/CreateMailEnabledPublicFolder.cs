// Demonstrates creating a public folder and making it able to receive e-mail.
//
// CreatePublicFolder adds a folder to the organisation's public folder tree with the
// given permissions. MailEnablePublicFolder gives it an e-mail address, so messages can
// be sent to the folder like to a mailbox; MailDisablePublicFolder removes the address
// again. This needs public folders to be set up and the rights to manage them.
// The example uses a uniquely named folder and deletes it at the end.

using System;
using Aspose.Email.Clients.Exchange;

namespace Aspose.Email.Examples.EWS
{
    internal static class CreateMailEnabledPublicFolder
    {
        public static void Run()
        {
            var folderName = "Aspose-" + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                // Everyone in the organisation may read; nobody else.
                var permissions = new ExchangeFolderPermissionCollection
                {
                    new ExchangeFolderPermission(ExchangeFolderUserInfo.DefaultUser)
                    {
                        PermissionLevel = ExchangeFolderPermissionLevel.Reviewer
                    }
                };

                var folder = client.CreatePublicFolder(folderName, permissions);
                Console.WriteLine($"Created public folder '{folder.DisplayName}'.");

                try
                {
                    client.MailEnablePublicFolder(folder.Uri);
                    Console.WriteLine("Mail-enabled it: it can now receive e-mail.");

                    client.MailDisablePublicFolder(folder.Uri);
                    Console.WriteLine("Mail-disabled it again.");
                }
                finally
                {
                    client.DeleteFolder(folder.Uri, true);
                    Console.WriteLine($"Deleted '{folder.DisplayName}'.");
                }
            }
        }
    }
}
