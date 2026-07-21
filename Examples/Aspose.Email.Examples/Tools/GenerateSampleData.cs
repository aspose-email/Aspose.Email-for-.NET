// Regenerates the sample input files under Data\ that are produced with Aspose.Email
// itself rather than captured from a real mail client.
//
// The generated files are committed to the repository, so this is not run as part of a
// normal example run: fixtures have to stay inert, otherwise a broken change would just
// rewrite them and still look green. Run this only when a fixture needs to be recreated,
// then copy the files it writes from Out\ into Data\ as printed below.

using System;
using System.IO;
using System.Text;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.Tools
{
    internal static class GenerateSampleData
    {
        public static void Run()
        {
            var sep = Path.DirectorySeparatorChar;

            Console.WriteLine("Generated sample input files:");
            Report(CreateMessageWithReceiptTimes(), $"Data{sep}Outlook{sep}TestMessage.msg");
            Report(CreateEmlWithEmbeddedMsg(), $"Data{sep}Email{sep}sample.eml");
            Report(CreateMessageWithVoteResults(), $"Data{sep}Outlook{sep}VoteResults.msg");
        }

        private static void Report(string generated, string destination)
        {
            Console.WriteLine($"  {generated}");
            Console.WriteLine($"      -> copy to {destination}");
        }

        // Builds Data\Outlook\TestMessage.msg, used by RetrieveReadAndDeliveryReceiptInformation.
        //
        // Delivery and read receipts are recorded per recipient, as PT_SYSTIME properties on
        // the recipient rather than on the message, which is why they are set in this loop.
        private static string CreateMessageWithReceiptTimes()
        {
            var path = Data.Out/"TestMessage.msg";

            using (var msg = new MapiMessage(
                       "sender@example.com",
                       "recipient@example.com",
                       "Read and delivery receipt sample",
                       "A message carrying per-recipient delivery and read receipt times."))
            {
                msg.Recipients.Add("second.recipient@example.com", "Second Recipient", MapiRecipientType.MAPI_TO);

                var deliveredAt = new DateTime(2024, 1, 15, 9, 30, 0, DateTimeKind.Utc);
                var readAt = new DateTime(2024, 1, 15, 10, 5, 0, DateTimeKind.Utc);

                foreach (var recipient in msg.Recipients)
                {
                    recipient.SetProperty(MapiProperty.CreateMapiPropertyFromDateTime(
                        MapiPropertyTag.PR_RECIPIENT_TRACKSTATUS_TIME_DELIVERY, deliveredAt));

                    recipient.SetProperty(MapiProperty.CreateMapiPropertyFromDateTime(
                        MapiPropertyTag.PR_RECIPIENT_TRACKSTATUS_TIME_READ, readAt));

                    // Stagger the second recipient so the sample shows differing times.
                    deliveredAt = deliveredAt.AddMinutes(7);
                    readAt = readAt.AddMinutes(21);
                }

                msg.Save(path);
            }

            return path;
        }

        // Builds Data\Outlook\VoteResults.msg, used by ReadVoteResultsInformation.
        //
        // The sample messages that carry voting buttons have no responses recorded, so
        // there is nothing for that example to report. Here each recipient has actually
        // voted: a response string, a response time, and a track status.
        private static string CreateMessageWithVoteResults()
        {
            var path = Data.Out/"VoteResults.msg";

            using (var msg = new MapiMessage(
                       "organiser@example.com",
                       "first.voter@example.com",
                       "Team lunch on Friday?",
                       "Please pick one of the options above."))
            {
                msg.Recipients.Add("second.voter@example.com", "Second Voter", MapiRecipientType.MAPI_TO);

                FollowUpManager.AddVotingButton(msg, "Yes");
                FollowUpManager.AddVotingButton(msg, "No");

                var responses = new[] { "Yes", "No" };
                var trackStatuses = new[] { MapiRecipientTrackStatus.Accepted, MapiRecipientTrackStatus.Declined };
                var votedAt = new DateTime(2024, 3, 4, 12, 0, 0, DateTimeKind.Utc);

                for (var i = 0; i < msg.Recipients.Count; i++)
                {
                    var recipient = msg.Recipients[i];

                    recipient.SetProperty(new MapiProperty(
                        MapiPropertyTag.PR_RECIPIENT_AUTORESPONSE_PROP_RESPONSE,
                        Encoding.Unicode.GetBytes(responses[i])));

                    recipient.SetProperty(MapiProperty.CreateMapiPropertyFromDateTime(
                        MapiPropertyTag.PR_RECIPIENT_TRACKSTATUS_TIME, votedAt.AddMinutes(i * 17)));

                    recipient.RecipientTrackStatus = trackStatuses[i];
                }

                msg.Save(path);
            }

            return path;
        }

        // Builds Data\Email\sample.eml, used by PreservingEmbeddedMsgFormat.
        //
        // The embedded item has to be a genuine MSG rather than a nested rfc822 part,
        // otherwise there is no MSG format left to preserve. Attaching MSG bytes with
        // Attachment(stream, "embedded.msg") is not enough - that produces a plain file
        // attachment. Adding a MapiMessage to a MapiMessage creates a real embedded
        // message, and TNEF encoding is what carries it through the EML intact.
        private static string CreateEmlWithEmbeddedMsg()
        {
            var path = Data.Out/"sample.eml";

            using (var embedded = new MapiMessage(
                       "inner.sender@example.com",
                       "inner.recipient@example.com",
                       "Embedded Outlook message",
                       "This message travels inside the EML as an embedded Outlook message."))
            using (var outer = new MapiMessage(
                       "sender@example.com",
                       "recipient@example.com",
                       "Message with an embedded MSG attachment",
                       "The attachment of this message is an Outlook message."))
            {
                outer.Attachments.Add("embedded.msg", embedded);

                var eml = outer.ToMailMessage(new MailConversionOptions { ConvertAsTnef = true });
                eml.Save(path, SaveOptions.DefaultEml);
            }

            return path;
        }
    }
}
