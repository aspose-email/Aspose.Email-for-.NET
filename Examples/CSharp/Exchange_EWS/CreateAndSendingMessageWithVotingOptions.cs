using System;
using System.Net;
using Aspose.Email.Mime;
using Aspose.Email.Mapi;
using Aspose.Email.Clients.Exchange.WebService;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Exchange_EWS
{
    class CreateAndSendingMessageWithVotingOptions
    {
        public static void Run()
        {
            // ExStart:CreateAndSendingMessageWithVotingOptions
            string address = "firstname.lastname@aspose.com";

            IEWSClient client = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain");
            MailMessage message = CreateTestMessage(address);

            // Set FollowUpOptions Buttons
            FollowUpOptions options = new FollowUpOptions();
            options.VotingButtons = "Yes;No;Maybe;Exactly!";

            client.Send(message, options);
            // ExEnd:CreateAndSendingMessageWithVotingOptions
        }

        // ExStart:CreateTestMessage
        private static MailMessage CreateTestMessage(string address)
        {
            MailMessage eml = new MailMessage(
                address,
                address,
                "Flagged message",
                "Make it nice and short, but descriptive. The description may appear in search engines' search results pages...");

            return eml;
        }
        // ExEnd:CreateTestMessage
    }
}