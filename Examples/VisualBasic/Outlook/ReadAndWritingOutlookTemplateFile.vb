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
    Class ReadAndWritingOutlookTemplateFile
        Public Shared Sub Run()
            'ExStart:ReadAndWritingOutlookTemplateFile
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()

            ' Load the Outlook template (OFT) file in MailMessage's instance
            Dim message As MailMessage = MailMessage.Load(dataDir & Convert.ToString("sample.oft"), New MsgLoadOptions())

            ' Set the sender and recipients information
            Dim senderDisplayName As String = "John"
            Dim senderEmailAddress As String = "john@abc.com"
            Dim recipientDisplayName As String = "William"
            Dim recipientEmailAddress As String = "william@xzy.com"

            message.Sender = New MailAddress(senderEmailAddress, senderDisplayName)
            message.[To].Add(New MailAddress(recipientEmailAddress, recipientDisplayName))
            message.HtmlBody = message.HtmlBody.Replace("DisplayName", (Convert.ToString("<b>") & recipientDisplayName) + "</b>")

            ' Set the name, location and time in email body
            Dim meetingLocation As String = "<u>" + "Hall 1, Convention Center, New York, USA" + "</u>"
            Dim meetingTime As String = "<u>" + "Monday, June 28, 2010" + "</u>"
            message.HtmlBody = message.HtmlBody.Replace("MeetingPlace", meetingLocation)
            message.HtmlBody = message.HtmlBody.Replace("MeetingTime", meetingTime)

            ' Save the message in MSG format and open in Office Outlook
            Dim msg As MapiMessage = MapiMessage.FromMailMessage(message)
            msg.SetMessageFlags(MapiMessageFlags.MSGFLAG_UNSENT)
            msg.Save(dataDir & Convert.ToString("ReadAndWritingOutlookTemplateFile_out.msg"))
            'ExEnd:ReadAndWritingOutlookTemplateFile
        End Sub
    End Class
End Namespace
