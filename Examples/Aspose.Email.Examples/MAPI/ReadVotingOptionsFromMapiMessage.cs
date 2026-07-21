// Demonstrates how to set voting buttons on a MapiMessage and read them back.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class ReadVotingOptionsFromMapiMessage
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

            // Reading them back gives an array, one entry per button.
            var saved = MapiMessage.Load(outputPath);
            var buttons = FollowUpManager.GetVotingButtons(saved);

            Console.WriteLine($"Poll: {saved.Subject}");
            Console.WriteLine($"Read back {buttons.Length} voting button(s):");
            foreach (var button in buttons)
                Console.WriteLine($"    {button}");

            Console.WriteLine($"\nSaved to {outputPath}");
        }
    }
}
