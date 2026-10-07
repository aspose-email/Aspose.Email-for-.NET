// Demonstrates the folder attributes the LIST command reports.
//
// Besides the name, every folder in a LIST response carries attributes: whether it can
// be selected, whether it has subfolders (CHILDREN, RFC 3348) or can never have any,
// and whether the server marked it as interesting. Servers with LIST-EXTENDED
// (RFC 5258) report more on request - see ListFoldersWithExtendedOptions.

using System;

namespace Aspose.Email.Examples.IMAP
{
    internal static class IMAP4ExtendedListCommand
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                var folders = client.ListFolders();

                Console.WriteLine($"LIST-EXTENDED supported: {client.ExtendedListSupported}");
                Console.WriteLine($"CHILDREN supported:      {client.ChildrenSupported}");
                Console.WriteLine($"\n{folders.Count} folder(s):");

                foreach (var folder in folders)
                {
                    Console.WriteLine($"  {folder.Name}");
                    Console.WriteLine($"    selectable: {folder.Selectable}, has children: {folder.HasChildren}, " +
                                      $"no inferiors: {folder.NoInferiors}, marked: {folder.Marked}");
                }
            }
        }
    }
}
