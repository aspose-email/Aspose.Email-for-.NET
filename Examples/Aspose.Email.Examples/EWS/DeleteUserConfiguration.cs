// Demonstrates removing a user configuration (settings stored in the mailbox).
//
// DeleteUserConfiguration takes the configuration's name and folder. Reading it
// afterwards fails, which the example shows. It creates the configuration on a
// uniquely named folder first and deletes the folder at the end.

using System;
using Aspose.Email.Clients.Exchange.WebService;

namespace Aspose.Email.Examples.EWS
{
    internal static class DeleteUserConfiguration
    {
        public static void Run()
        {
            var folderName = "Aspose-" + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var folder = client.CreateFolder(client.MailboxInfo.InboxUri, folderName);
                var name = new UserConfigurationName("Aspose.Example.Settings", folder.Uri);

                try
                {
                    var created = new UserConfiguration(name);
                    created.Dictionary.Add("theme", "dark");
                    client.CreateUserConfiguration(created);

                    client.DeleteUserConfiguration(name);
                    Console.WriteLine($"Deleted the configuration '{name.Name}'.");

                    try
                    {
                        client.GetUserConfiguration(name);
                        Console.WriteLine("It can still be read.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Reading it now fails as expected: {ex.Message}");
                    }
                }
                finally
                {
                    client.DeleteFolder(folder.Uri, true);
                }
            }
        }
    }
}
