Imports Aspose.Email.Outlook

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'


Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class CreatePollUsingMapiMessage
        Public Shared Sub Run()
            ' ExStart:CreatePollUsingMapiMessage
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()

            Dim msg As MapiMessage = CreateTestMessage(False)

            Dim options As New FollowUpOptions()
            options.VotingButtons = "YesNoMaybeExactly!"

            FollowUpManager.SetOptions(msg, options)

            msg.Save(dataDir & Convert.ToString("MapiMsgWithPoll.msg"))
            ' ExEnd:CreatePollUsingMapiMessage
        End Sub
        Private Shared Function CreateTestMessage(draft As Boolean) As MapiMessage
            Dim msg As New MapiMessage("from@test.com", "to@test.com", "Flagged message", "Make it nice and short, but descriptive. The description may appear in search engines' search results pages...")

            If Not draft Then
                msg.SetMessageFlags(msg.Flags Xor MapiMessageFlags.MSGFLAG_UNSENT)
            End If

            Return msg
        End Function
    End Class
End Namespace
