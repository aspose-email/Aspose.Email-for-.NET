// Demonstrates telling ordinary mail from other item types in a listing.
//
// A mail folder can hold more than messages: meeting requests and responses, delivery
// reports, posts. MessageInfoType and MessageClass on ExchangeMessageInfo say what each
// item is, so it can be handled accordingly before anything is downloaded.

using System;

namespace Aspose.Email.Examples.EWS
{
    internal static class GetMessageTypeFromExchangeMessageInfo
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                foreach (var info in client.ListMessages(client.MailboxInfo.InboxUri, 20))
                    Console.WriteLine($"  {info.MessageInfoType,-14} {info.MessageClass,-28} {info.Subject}");
            }
        }
    }
}
