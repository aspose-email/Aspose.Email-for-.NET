// Demonstrates how to read the responses recipients gave to a poll, together with
// the time each of them voted.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class ReadVoteResultsInformation
    {
        public static void Run()
        {
            var msg = MapiMessage.Load(Data.Mapi/"VoteResults.msg");

            Console.WriteLine($"Poll: {msg.Subject}");
            Console.WriteLine($"Options: {string.Join(", ", FollowUpManager.GetVotingButtons(msg))}\n");

            foreach (var recipient in msg.Recipients)
            {
                Console.WriteLine($"Recipient:     {recipient.DisplayName}");
                Console.WriteLine($"Response:      {GetResponse(recipient)}");
                Console.WriteLine($"Response time: {GetResponseTime(recipient)}");
                Console.WriteLine($"Track status:  {recipient.RecipientTrackStatus}");
                Console.WriteLine();
            }
        }

        // These properties only exist once a recipient has actually voted, so the
        // indexer returns null for anyone who has not responded yet.
        private static string GetResponse(MapiRecipient recipient)
        {
            var property = recipient.Properties[MapiPropertyTag.PR_RECIPIENT_AUTORESPONSE_PROP_RESPONSE];
            return property == null ? "(no response yet)" : property.GetString();
        }

        private static string GetResponseTime(MapiRecipient recipient)
        {
            var property = recipient.Properties[MapiPropertyTag.PR_RECIPIENT_TRACKSTATUS_TIME];
            return property == null ? "(no response yet)" : property.GetDateTime().ToString("u");
        }
    }
}
