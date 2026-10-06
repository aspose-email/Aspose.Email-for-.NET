// Demonstrates how to create a folder and a subfolder, and how to check that a folder
// exists and get its details in the same call.
//
// Nested names are joined with the server's hierarchy delimiter - '/' on most servers,
// '.' on some - which ImapClient.Delimiter reports once the client has listed folders.
// The folders get a unique name and are deleted again at the end.

using System;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class CreateFolders
    {
        public static void Run()
        {
            var parentName = "Aspose-" + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                if (!client.ExistFolder(parentName))
                    client.CreateFolder(parentName);

                var childName = parentName + client.Delimiter + "Reports";
                client.CreateFolder(childName);

                try
                {
                    Console.WriteLine($"Hierarchy delimiter: '{client.Delimiter}'");

                    ImapFolderInfo info;
                    if (client.ExistFolder(childName, out info))
                    {
                        Console.WriteLine($"Created '{info.Name}'");
                        Console.WriteLine($"  selectable:    {info.Selectable}");
                        Console.WriteLine($"  messages:      {info.TotalMessageCount}");
                        Console.WriteLine($"  can have kids: {!info.NoInferiors}");
                    }

                    Console.WriteLine($"\nSubfolders of '{parentName}':");
                    foreach (var folder in client.ListFolders(parentName))
                        Console.WriteLine("  " + folder.Name);
                }
                finally
                {
                    // Children first: some servers refuse to delete a folder that has any.
                    client.DeleteFolder(childName);
                    client.DeleteFolder(parentName);
                }
            }
        }
    }
}
