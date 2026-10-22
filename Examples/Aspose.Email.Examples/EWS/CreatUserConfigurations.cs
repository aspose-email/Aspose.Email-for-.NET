// Demonstrates storing application settings in the mailbox (user configuration).
//
// A user configuration is a named settings object attached to a folder. It is hidden
// from the user and travels with the mailbox, which makes it a place for an add-in or
// application to keep per-user settings. Its Dictionary holds key/value pairs; XmlData
// and BinaryData hold larger payloads. The example attaches a configuration to a
// uniquely named folder and deletes the folder, and with it the configuration, at the
// end.

using System;
using Aspose.Email.Clients.Exchange.WebService;

namespace Aspose.Email.Examples.EWS
{
    internal static class CreatUserConfigurations
    {
        public static void Run()
        {
            var folderName = "Aspose-" + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var folder = client.CreateFolder(client.MailboxInfo.InboxUri, folderName);

                try
                {
                    var name = new UserConfigurationName("Aspose.Example.Settings", folder.Uri);
                    var configuration = new UserConfiguration(name);
                    configuration.Dictionary.Add("theme", "dark");
                    configuration.Dictionary.Add("pageSize", "50");

                    client.CreateUserConfiguration(configuration);
                    Console.WriteLine($"Created the configuration '{configuration.UserConfigurationName.Name}' " +
                                      $"with {configuration.Dictionary.Count} setting(s).");
                }
                finally
                {
                    client.DeleteFolder(folder.Uri, true);
                }
            }
        }
    }
}
