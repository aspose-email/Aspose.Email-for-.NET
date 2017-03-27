using Aspose.Email.Mapi;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class DeleteVotingButtonFromMessage
    {
        public static void Run()
        {
            // ExStart:DeletVotingButtonFromMessage
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Outlook();

            // Create New Message and set FollowUpOptions, FollowUpManager properties
            MapiMessage msg = CreateTestMessage(false);

            FollowUpOptions options = new FollowUpOptions();
            options.VotingButtons = "Yes;No;Maybe;Exactly!";
            FollowUpManager.SetOptions(msg, options);
            msg.Save(dataDir + "MapiMsgWithPoll.msg");
            FollowUpManager.RemoveVotingButton(msg, "Exactly!"); // Deleting a single button OR
            FollowUpManager.ClearVotingButtons(msg); // Deleting all buttons from a MapiMessage
            msg.Save(dataDir + "MapiMsgWithPoll.msg");
            // ExEnd:DeletVotingButtonFromMessage
        }

        private static MapiMessage CreateTestMessage(bool draft)
        {
            // ExEnd:CreateTestMessage
            MapiMessage msg = new MapiMessage(
            "from@test.com",
            "to@test.com",
            "Flagged message",
            "Make it nice and short, but descriptive. The description may appear in search engines' search results pages...");

            if (!draft)
            {
                msg.SetMessageFlags(msg.Flags ^ MapiMessageFlags.MSGFLAG_UNSENT);
            }
            return msg;
            // ExEnd:CreateTestMessage
        }
    }
}