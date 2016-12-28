Imports Aspose.Email.Outlook
Imports Aspose.Email.Mail

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class SetFollowUpflag
        Public Shared Sub Run()
            'ExStart:SetFollowUpflag
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()

            Dim mailMsg As New MailMessage()
            mailMsg.Sender = "AETest12@gmail.com"
            mailMsg.[To] = "receiver@gmail.com"
            mailMsg.Body = "This message will test if follow up options can be added to a new mapi message."
            Dim mapi As MapiMessage = MapiMessage.FromMailMessage(mailMsg)

            Dim dtStartDate As New DateTime(2013, 5, 23, 14, 40, 0)
            Dim dtReminderDate As New DateTime(2013, 5, 23, 16, 40, 0)
            Dim dtDueDate As DateTime = dtReminderDate.AddDays(1)

            Dim options As New FollowUpOptions("Follow Up", dtStartDate, dtDueDate, dtReminderDate)
            FollowUpManager.SetOptions(mapi, options)
            mapi.Save(dataDir & Convert.ToString("SetFollowUpflag_out.msg"))
            'ExEnd:SetFollowUpflag
        End Sub
    End Class
End Namespace