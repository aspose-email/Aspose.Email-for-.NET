Imports Aspose.Email.Outlook

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class GetTheTextAndRTFBodies
        Public Shared Sub Run()
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()

            ' Load mail message
            Dim msg As MapiMessage = MapiMessage.FromMailMessage(dataDir & Convert.ToString("Message.eml"))

            Dim itemBase As MapiMessageItemBase = New MapiMessage()
            ' Text body
            If itemBase.Body IsNot Nothing Then
                Console.WriteLine(msg.Body)
            Else
                Console.WriteLine("There's no text body.")
            End If

            ' RTF body
            If itemBase.BodyRtf IsNot Nothing Then
                Console.WriteLine(msg.BodyRtf)
            Else
                Console.WriteLine("There's no RTF body.")
            End If
        End Sub

    End Class
End Namespace