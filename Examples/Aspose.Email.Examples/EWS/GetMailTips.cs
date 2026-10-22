// Demonstrates mail tips - what Outlook warns about before you send.
//
// GetMailTips checks recipients before a message is sent: whether they are out of
// office (with their automatic reply), whether the address is invalid or the mailbox
// full, how large a message they accept, and how many members a group has. Put the
// addresses you are about to write to into the list below.

using System;
using Aspose.Email.Clients.Exchange.WebService;

namespace Aspose.Email.Examples.EWS
{
    internal static class GetMailTips
    {
        public static void Run()
        {
            var sender = ClientBuilder.EwsUserName;
            var recipients = new MailAddressCollection
            {
                new MailAddress(sender, true),
                new MailAddress("no.such.user@example.com", true)
            };

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var options = new GetMailTipsOptions(new MailAddress(sender, true), recipients, MailTipsType.All);

                foreach (var tip in client.GetMailTips(options))
                {
                    Console.WriteLine(tip.RecipientAddress);
                    Console.WriteLine($"  invalid recipient:  {tip.InvalidRecipient}");
                    Console.WriteLine($"  mailbox full:       {tip.MailboxFull}");
                    Console.WriteLine($"  max message size:   {tip.MaxMsgSize} bytes");
                    if (tip.OutOfOffice?.ReplyBody != null)
                        Console.WriteLine($"  out of office:      {tip.OutOfOffice.ReplyBody.Message}");
                }
            }
        }
    }
}
