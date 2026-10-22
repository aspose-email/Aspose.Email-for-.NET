// Demonstrates sending a message with Outlook voting buttons.
//
// FollowUpOptions.VotingButtons takes the button captions separated by semicolons.
// Outlook shows them above the message, and each vote comes back to the sender as a
// reply that Outlook tallies. The message goes to the configured mailbox itself.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.EWS
{
    internal static class CreateAndSendingMessageWithVotingOptions
    {
        public static void Run()
        {
            var self = ClientBuilder.EwsUserName;

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var message = new MailMessage(self, self, "Team lunch on Friday?", "Please vote below.");
                var options = new FollowUpOptions { VotingButtons = "Yes;No;Maybe" };

                client.Send(message, options);
                Console.WriteLine($"Sent '{message.Subject}' with the buttons {options.VotingButtons} to {self}.");
            }
        }
    }
}
