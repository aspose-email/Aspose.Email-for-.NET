Imports Aspose.Email.Outlook

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email
    Class IdentifyInlineAndRegularAttachments
        Public Shared Sub Run()
            ' ExStart:IdentifyInlineAndRegularAttachments           
            Dim dataDir As String = RunExamples.GetDataDir_Email()
            Dim message = MapiMessage.FromFile(dataDir & Convert.ToString("RemoveAttachments.msg"))
            Dim attachments = message.Attachments
            Dim count = attachments.Count

            Console.WriteLine("Total attachments " + count.ToString())
            For i As Integer = 0 To attachments.Count - 1
                Dim attachment = attachments(i)
                If IsInlineAttachment(attachment, message) Then
                    Console.WriteLine(attachment.LongFileName.ToString() + " is inline attachment")
                Else
                    Console.WriteLine(attachment.LongFileName.ToString() + " is regular attachment")
                End If
            Next

        End Sub

        Public Shared Function IsInlineAttachment(att As MapiAttachment, msg As MapiMessage) As Boolean
            Select Case msg.BodyType
                Case BodyContentType.PlainText
                    ' ignore indications for plain text messages
                    Return False

                Case BodyContentType.Html

                    ' check the PidTagAttachFlags property
                    If att.Properties.Contains(&H37140003) Then
                        Dim attachFlagsValue As System.Nullable(Of Long) = att.GetPropertyLong(&H37140003)
                        If attachFlagsValue IsNot Nothing AndAlso (attachFlagsValue And &H4) = &H4 Then
                            ' check PidTagAttachContentId property
                            If att.Properties.Contains(MapiPropertyTag.PR_ATTACH_CONTENT_ID) OrElse att.Properties.Contains(MapiPropertyTag.PR_ATTACH_CONTENT_ID_W) Then
                                Dim contentId As String = If(att.Properties.Contains(MapiPropertyTag.PR_ATTACH_CONTENT_ID), att.Properties(MapiPropertyTag.PR_ATTACH_CONTENT_ID).GetString(), att.Properties(MapiPropertyTag.PR_ATTACH_CONTENT_ID_W).GetString())
                                If msg.BodyHtml.Contains(contentId) Then
                                    Return True
                                End If
                            End If
                            ' check PidTagAttachContentLocation property
                            If att.Properties.Contains(&H3713001E) OrElse att.Properties.Contains(&H3713001F) Then
                                Return True
                            End If
                        ElseIf (att.Properties.Contains(&H3716001F) AndAlso att.GetPropertyString(&H3716001F) = "inline") OrElse (att.Properties.Contains(&H3716001E) AndAlso att.GetPropertyString(&H3716001E) = "inline") Then
                            Return True
                        End If
                    ElseIf (att.Properties.Contains(&H3716001F) AndAlso att.GetPropertyString(&H3716001F) = "inline") OrElse (att.Properties.Contains(&H3716001E) AndAlso att.GetPropertyString(&H3716001E) = "inline") Then
                        Return True
                    End If
                    Return False

                Case BodyContentType.Rtf

                    ' If the body is RTF, then all OLE attachments are inline attachments.
                    ' OLE attachments have 0x00000006 for the value of the PidTagAttachMethod property
                    If att.Properties.Contains(MapiPropertyTag.PR_ATTACH_METHOD) Then
                        Return att.GetPropertyLong(MapiPropertyTag.PR_ATTACH_METHOD) = &H6
                    End If
                    Return False
                Case Else
                    Throw New ArgumentOutOfRangeException()
            End Select
        End Function
        ' ExEnd:IdentifyInlineAndRegularAttachments
    End Class
End Namespace
