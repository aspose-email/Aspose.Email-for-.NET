// Demonstrates how to read the voting buttons of a message through FollowUpOptions,
// which also exposes the other follow-up settings at the same time.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class ReadingVotingOptions
    {
        public static void Run()
        {
            var message = MapiMessage.Load(Data.Mapi/"MessageWithVotingButtons.msg");

            // Reading the whole options object is useful when more than the buttons is
            // needed - it also carries the categories and the flag settings.
            var options = FollowUpManager.GetOptions(message);

            // The buttons come back as one string, separated by semicolons.
            var votingButtons = options.VotingButtons;

            Console.WriteLine($"Subject:        {message.Subject}");
            Console.WriteLine($"Voting buttons: {(string.IsNullOrEmpty(votingButtons) ? "(none)" : votingButtons)}");
            Console.WriteLine($"Categories:     {(string.IsNullOrEmpty(options.Categories) ? "(none)" : options.Categories)}");
        }
    }
}
