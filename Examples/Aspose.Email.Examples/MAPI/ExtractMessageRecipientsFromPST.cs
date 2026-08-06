// Demonstrates how to read just the recipients of PST messages, which is much
// cheaper than extracting every message in full.

using System;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.MAPI
{
    internal static class ExtractMessageRecipientsFromPST
    {
        public static void Run()
        {
            using (var pst = PersonalStorage.FromFile(Data.Mapi/"Sub.pst", false))
            {
                var inbox = pst.RootFolder.GetSubFolder("Inbox");

                foreach (var messageInfo in inbox.EnumerateMessages())
                {
                    Console.WriteLine($"Message: {messageInfo.Subject}");

                    // The MessageInfo overload is a convenience wrapper around the
                    // entry id one - ExtractRecipients(messageInfo.EntryIdString).
                    var recipients = pst.ExtractRecipients(messageInfo);

                    foreach (var recipient in recipients)
                        Console.WriteLine($"  {recipient.RecipientType,-8} {recipient.DisplayName} <{recipient.EmailAddress}>");
                }
            }
        }
    }
}
