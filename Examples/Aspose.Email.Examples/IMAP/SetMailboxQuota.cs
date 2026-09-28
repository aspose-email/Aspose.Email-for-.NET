// Demonstrates the SETQUOTA command, which changes the limit of a quota root.
//
// Setting quotas is an administrative operation: most servers only accept it from an
// admin account, and many hosted services do not accept it at all. Expect a NO response
// for an ordinary mailbox; the example reports it instead of failing.

using System;
using System.Linq;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class SetMailboxQuota
    {
        public static void Run()
        {
            const int storageLimitKiB = 2 * 1024 * 1024;  // 2 GiB

            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.GetCapabilities();

                if (!client.QuotaSupported)
                {
                    Console.WriteLine("The server does not support QUOTA.");
                    return;
                }

                var root = client.GetQuotaRoot(ImapFolderInfo.InBox).FirstOrDefault();
                if (root == null)
                {
                    Console.WriteLine("The Inbox is not under any quota root.");
                    return;
                }

                try
                {
                    var quotas = client.SetQuota(root.QuotaRootName, "STORAGE", storageLimitKiB);

                    Console.WriteLine($"New limits of quota root '{root.QuotaRootName}':");
                    foreach (var quota in quotas)
                        Console.WriteLine($"  {quota.ResourceName,-8} {quota.Used} of {quota.Limit}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"The server refused to change the quota: {ex.Message}");
                }
            }
        }
    }
}
