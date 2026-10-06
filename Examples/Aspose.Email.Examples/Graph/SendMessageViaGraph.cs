// Demonstrates the ways a message can be sent through Microsoft Graph.
//
// Send takes a message and posts it in one call. SendAsMime posts the raw MIME instead,
// which is what preserves a message that was built or signed elsewhere. Creating a draft
// first and sending it by id is the third route, and the one to use when the message has
// to be reviewed or added to before it goes out.

using System;
using System.Linq;
using Aspose.Email.Clients.Graph;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.Graph
{
    internal static class SendMessageViaGraph
    {
        public static void Run()
        {
            if (!ClientBuilder.IsGraphConfigured)
            {
                GraphExampleInfo.PrintNotConfigured();
                return;
            }

            var from = ClientBuilder.GraphMailboxId;
            const string to = "recipient@example.com";
            const string draftSubject = "Drafted through Graph";

            using (var client = ClientBuilder.Graph(AuthType.ModernWithAppPermission))
            {
                // A MailMessage goes straight out.
                var eml = new MailMessage(from, to, "Sent through Graph", "Plain MIME message.");
                client.Send(eml);
                Console.WriteLine("Sent a MailMessage.");

                // A MapiMessage can be sent too, with control over the Sent Items copy.
                var msg = new MapiMessage(from, to, "Sent through Graph as MSG", "Outlook message.",
                    OutlookMessageFormat.Unicode);
                client.Send(msg, true);
                Console.WriteLine("Sent a MapiMessage, keeping a copy in Sent Items.");

                // SendAsMime posts the MIME as-is, which keeps a signature intact.
                client.SendAsMime(msg);
                Console.WriteLine("Sent the same message as raw MIME.");

                // Draft first, send later: useful when something has to be added to the
                // message between creating and sending it.
                var draft = new MapiMessage(from, to, draftSubject, "Draft body.",
                    OutlookMessageFormat.Unicode);

                client.CreateMessage(KnownFolders.Drafts, draft);
                Console.WriteLine($"Created a draft: {draftSubject}");

                // Send takes the Graph item id, which comes from a listing rather than
                // from the message object.
                var created = client.ListMessages(KnownFolders.Drafts, null)
                    .Cast<MessageInfo>()
                    .FirstOrDefault(m => m.Subject == draftSubject);

                if (created == null)
                {
                    Console.WriteLine("The draft was not found in the Drafts folder.");
                    return;
                }

                client.Send(created.ItemId);
                Console.WriteLine($"Sent the draft by its item id ({created.ItemId}).");
            }
        }
    }
}
