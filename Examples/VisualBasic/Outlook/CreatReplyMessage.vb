Imports Aspose.Email.Outlook

' This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET 
'   API reference when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq 
'   for more information. If you do not wish to use NuGet, you can manually download 
'   Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'   install it and then add its reference to this project. For any issues, questions or suggestions 
'   please feel free to contact us using http://www.aspose.com/community/forums/default.aspx            
'

Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class CreatReplyMessage
        Public Shared Sub Run()
            ' ExStart:CreatReplyMessage
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()

            Dim originalMsg As MapiMessage = MapiMessage.FromFile(dataDir & Convert.ToString("message1.msg"))
            Dim builder As New ReplyMessageBuilder()

            ' Set ReplyMessageBuilder Properties
            builder.ReplyAll = True
            builder.AdditionMode = OriginalMessageAdditionMode.Textpart
            builder.ResponseText = "<p><b>Dear Friend,</b></p> I want to do is introduce my co-author and co-teacher. <p><a href=""www.google.com"">This is a first link</a></p><p><a href=""www.google.com"">This is a second link</a></p>"

            Dim replyMsg As MapiMessage = builder.BuildResponse(originalMsg)
            replyMsg.Save(dataDir & Convert.ToString("reply_out.msg"))
            ' ExEnd:CreatReplyMessage
        End Sub
    End Class
End Namespace