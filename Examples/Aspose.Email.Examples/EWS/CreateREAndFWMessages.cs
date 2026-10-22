// Demonstrates replying to and forwarding a message that is on the server.
//
// Reply, ReplyAll and Forward take the new message and the ExchangeMessageInfo of the
// original; Exchange builds the "RE:"/"FW:" message with the original quoted and sends
// it. Here the original is a test message from the configured mailbox to itself, so
// the replies and the forward arrive in that mailbox too.

using System;
using System.Linq;

namespace Aspose.Email.Examples.EWS
{
    internal static class CreateREAndFWMessages
    {
        public static void Run()
        {
            var self = ClientBuilder.EwsUserName;
            var folderName = "Aspose-" + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var folder = client.CreateFolder(client.MailboxInfo.InboxUri, folderName);

                try
                {
                    client.AppendMessage(folder.Uri, new MailMessage(self, self, "Please reply", "Please reply to me."));
                    var original = client.ListMessages(folder.Uri).First();

                    client.Reply(new MailMessage(self, self, "", "Here is my reply."), original);
                    client.ReplyAll(new MailMessage(self, self, "", "Replying to everyone."), original);
                    client.Forward(new MailMessage(self, self, "", "Forwarding this to you."), original);

                    Console.WriteLine($"Sent a reply, a reply-all and a forward of '{original.Subject}' to {self}.");
                }
                finally
                {
                    client.DeleteFolder(folder.Uri, true);
                }
            }
        }
    }
}
