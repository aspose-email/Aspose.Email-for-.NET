// Demonstrates the NAMESPACE command (RFC 2342).
//
// A namespace tells you where folders live and how their names are built: the prefix to
// put in front of a new top-level folder, and the hierarchy delimiter. Personal
// namespaces hold the user's own folders; other-users and shared namespaces expose
// mailboxes the user has been given access to.

using System;

namespace Aspose.Email.Examples.IMAP
{
    internal static class ListMailboxNamespaces
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.GetCapabilities();

                if (!client.NamespaceSupported)
                {
                    Console.WriteLine("The server does not support NAMESPACE.");
                    return;
                }

                var namespaces = client.GetNamespaces();
                Console.WriteLine($"{namespaces.Length} namespace(s):");

                foreach (var ns in namespaces)
                {
                    Console.WriteLine($"  {ns.NamespaceType,-10}  prefix '{ns.Prefix}', " +
                                      $"delimiter '{ns.HierarchyDelimiter}'");
                }
            }
        }
    }
}
