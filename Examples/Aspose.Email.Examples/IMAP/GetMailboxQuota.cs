// Demonstrates how to read storage limits with the QUOTA extension (RFC 2087).
//
// A quota root is the unit a limit applies to, and several folders can share one.
// GetQuotaRoot tells you which roots a folder is counted against; GetQuota returns the
// usage and limit of a root per resource. STORAGE is measured in KiB, MESSAGE in
// messages.

using System;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class GetMailboxQuota
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.GetCapabilities();

                if (!client.QuotaSupported)
                {
                    Console.WriteLine("The server does not support QUOTA.");
                    return;
                }

                foreach (var root in client.GetQuotaRoot(ImapFolderInfo.InBox))
                {
                    Console.WriteLine($"'{root.MailboxName}' is counted against quota root '{root.QuotaRootName}':");

                    foreach (var quota in client.GetQuota(root.QuotaRootName))
                    {
                        var percent = quota.Limit > 0 ? 100.0 * quota.Used / quota.Limit : 0;
                        Console.WriteLine($"  {quota.ResourceName,-8} {quota.Used} of {quota.Limit} ({percent:F1}% used)");
                    }
                }
            }
        }
    }
}
