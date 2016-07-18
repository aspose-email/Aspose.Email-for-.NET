Imports Aspose.Email.Outlook

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'


Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class DeleteVotingButtonFromMessage
        Public Shared Sub Run()
            ' ExStart:DeletVotingButtonFromMessage
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()

            ' Create New Message and set FollowUpOptions, FollowUpManager properties
            Dim msg As MapiMessage = CreateTestMessage(False)

            Dim options As New FollowUpOptions()
            options.VotingButtons = "YesNoMaybeExactly!"
            FollowUpManager.SetOptions(msg, options)
            msg.Save(dataDir & Convert.ToString("MapiMsgWithPoll.msg"))
            FollowUpManager.RemoveVotingButton(msg, "Exactly!")
            ' Deleting a single button OR
            FollowUpManager.ClearVotingButtons(msg)
            ' Deleting all buttons from a MapiMessage
            msg.Save(dataDir & Convert.ToString("MapiMsgWithPoll.msg"))
            ' ExEnd:DeletVotingButtonFromMessage
        End Sub

        Private Shared Function CreateTestMessage(draft As Boolean) As MapiMessage
            ' ExEnd:CreateTestMessage
            Dim msg As New MapiMessage("from@test.com", "to@test.com", "Flagged message", "Make it nice and short, but descriptive. The description may appear in search engines' search results pages...")

            If Not draft Then
                msg.SetMessageFlags(msg.Flags Xor MapiMessageFlags.MSGFLAG_UNSENT)
            End If
            Return msg
            ' ExEnd:CreateTestMessage
        End Function
    End Class
End Namespace