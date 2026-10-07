// Demonstrates the life cycle of a folder: create it, rename it, delete it.
//
// RenameFolder takes the current and the new name; DeleteFolder removes the folder
// together with the messages in it. ExistFolder checks the outcome of each step.
// The example uses a uniquely named folder.

using System;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class DeletingAndRenamingFolders
    {
        public static void Run()
        {
            var suffix = Guid.NewGuid().ToString("N").Substring(0, 8);
            var originalName = "Aspose-Draft-" + suffix;
            var newName = "Aspose-Final-" + suffix;

            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.CreateFolder(originalName);
                Report(client, "Created", originalName, newName);

                client.RenameFolder(originalName, newName);
                Report(client, "Renamed", originalName, newName);

                client.DeleteFolder(newName);
                Report(client, "Deleted", originalName, newName);
            }
        }

        private static void Report(ImapClient client, string step, string originalName, string newName)
        {
            Console.WriteLine($"{step,-8} '{originalName}' exists: {client.ExistFolder(originalName),-5}  " +
                              $"'{newName}' exists: {client.ExistFolder(newName)}");
        }
    }
}
