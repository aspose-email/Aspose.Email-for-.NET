using System;
using System.Net;
using Aspose.Email.Mime;
using Aspose.Email.Mapi;
using Aspose.Email.Clients.Exchange.WebService;

namespace Aspose.Email.Examples.EWS
{
    class CreateAndSendingMessageWithVotingOptions
    {
        public static void Run()
        {
            var address = "you.address@domain";

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var message = CreateTestMessage(address);

                // Set FollowUpOptions Buttons
                var options = new FollowUpOptions
                {
                    VotingButtons = "Yes;No;Maybe;Exactly!"
                };

                client.Send(message, options);
            }
        }

        private static MailMessage CreateTestMessage(string address)
        {
            return new MailMessage(
                address,
                address,
                "Flagged message",
                "Make it nice and short, but descriptive.");
        }
    }
}