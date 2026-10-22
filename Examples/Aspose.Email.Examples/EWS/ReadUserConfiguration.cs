// Demonstrates reading a user configuration (settings stored in the mailbox).
//
// GetUserConfiguration takes the configuration's name and folder and returns it with its
// key/value pairs. The example first creates a configuration on a uniquely named folder,
// reads it back, and deletes the folder at the end.

using System;
using Aspose.Email.Clients.Exchange.WebService;

namespace Aspose.Email.Examples.EWS
{
    internal static class ReadUserConfiguration
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
                    created.Dictionary.Add("pageSize", "50");
                    client.CreateUserConfiguration(created);

                    var configuration = client.GetUserConfiguration(name);
                    Console.WriteLine($"Configuration '{configuration.UserConfigurationName.Name}' " +
                                      $"(id {configuration.Id}):");
                    foreach (var key in configuration.Dictionary.Keys)
                        Console.WriteLine($"  {key} = {configuration.Dictionary[key]}");
                }
                finally
                {
                    client.DeleteFolder(folder.Uri, true);
                }
            }
        }
    }
}
