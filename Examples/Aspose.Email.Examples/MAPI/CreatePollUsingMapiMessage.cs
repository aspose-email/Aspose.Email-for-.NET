// Demonstrates how to create a polling message with voting buttons using FollowUpManager.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class CreatePollUsingMapiMessage
    {
        public static void Run()
        {
            var msg = new MapiMessage(
                "from@test.com",
                "to@test.com",
                "Flagged message",
                "Make it nice and short, but descriptive. The description may appear in search engines' search results pages...");

            // Clear the draft flag, so Outlook treats the message as ready to send.
            msg.SetMessageFlags(msg.Flags ^ MapiMessageFlags.MSGFLAG_UNSENT);

            // The buttons are one string separated by semicolons.
            var options = new FollowUpOptions { VotingButtons = "Yes;No;Maybe;Exactly!" };
            FollowUpManager.SetOptions(msg, options);

            var outputPath = Data.Out/"MapiMsgWithPoll.msg";
            msg.Save(outputPath);

            Console.WriteLine($"Poll: {msg.Subject}");
            Console.WriteLine($"Options: {string.Join(", ", FollowUpManager.GetVotingButtons(msg))}");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}
