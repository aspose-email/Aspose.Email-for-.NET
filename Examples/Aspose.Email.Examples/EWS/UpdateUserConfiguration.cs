// Demonstrates changing a user configuration (settings stored in the mailbox).
//
// Read the configuration, change its Dictionary and pass it to UpdateUserConfiguration;
// the stored settings are replaced with the new ones. The example creates a
// configuration on a uniquely named folder, updates it and deletes the folder at the
// end.

using System;
using Aspose.Email.Clients.Exchange.WebService;

namespace Aspose.Email.Examples.EWS
{
    internal static class UpdateUserConfiguration
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

                    var configuration = client.GetUserConfiguration(name);
                    configuration.Dictionary["theme"] = "light";
                    configuration.Dictionary["language"] = "en";
                    client.UpdateUserConfiguration(configuration);

                    var updated = client.GetUserConfiguration(name);
                    Console.WriteLine("Settings after the update:");
                    foreach (var key in updated.Dictionary.Keys)
                        Console.WriteLine($"  {key} = {updated.Dictionary[key]}");
                }
                finally
                {
                    client.DeleteFolder(folder.Uri, true);
                }
            }
        }
    }
}
