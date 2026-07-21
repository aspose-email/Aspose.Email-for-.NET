// Demonstrates how to remove a single voting button and how to clear all voting buttons from a MapiMessage.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class DeleteVotingButtonFromMessage
    {
        public static void Run()
        {
            var msg = CreateTestMessage();

            var options = new FollowUpOptions { VotingButtons = "Yes;No;Maybe;Exactly!" };
            FollowUpManager.SetOptions(msg, options);
            msg.Save(Data.Out/"MapiMsgWithPoll.msg");
            Console.WriteLine($"All buttons:        {Describe(msg)}");

            // RemoveVotingButton takes one button away...
            FollowUpManager.RemoveVotingButton(msg, "Exactly!");
            Console.WriteLine($"After removing one: {Describe(msg)}");

            // ...while ClearVotingButtons takes the poll off the message entirely.
            FollowUpManager.ClearVotingButtons(msg);
            Console.WriteLine($"After clearing all: {Describe(msg)}");

            var outputPath = Data.Out/"MapiMsgWithPollCleared.msg";
            msg.Save(outputPath);
            Console.WriteLine($"\nSaved to {outputPath}");
        }

        private static string Describe(MapiMessage msg)
        {
            var buttons = FollowUpManager.GetVotingButtons(msg);
            return buttons.Length == 0 ? "(none)" : string.Join(", ", buttons);
        }

        private static MapiMessage CreateTestMessage()
        {
            var msg = new MapiMessage(
                "from@test.com",
                "to@test.com",
                "Flagged message",
                "Make it nice and short, but descriptive. The description may appear in search engines' search results pages...");

            msg.SetMessageFlags(msg.Flags ^ MapiMessageFlags.MSGFLAG_UNSENT);
            return msg;
        }
    }
}
