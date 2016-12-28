Imports Aspose.Email.Mail

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email
    Class PrintHeaderUsingMhtFormatOptions
        Public Shared Sub Run()
            ' ExStart:PrintHeaderUsingMhtFormatOptions
            ' The path to the File directory.
            Dim dataDir As String = RunExamples.GetDataDir_Email()
            Const pageHeader As String = "<div><div class='pageHeader'>&quot;Panditharatne, Mithra&quot; &lt;mithra.panditharatne@cibc.com&gt;<hr/></div>"
            Dim message As MailMessage = MailMessage.Load(dataDir & Convert.ToString("Message.eml"))
            Dim mailFormatter As New MhtMessageFormatter()
            Dim copyMessage As MailMessage = message.Clone()
            mailFormatter.Format(copyMessage)
            Console.WriteLine(If(copyMessage.HtmlBody.Contains(pageHeader), "True", "False"))
            Dim options As MhtFormatOptions = MhtFormatOptions.HideExtraPrintHeader Or MhtFormatOptions.WriteCompleteEmailAddress
            mailFormatter.Format(message, options)
            If Not message.HtmlBody.Contains(pageHeader) Then
                Console.WriteLine("True")
            Else
                Console.WriteLine("False")
            End If
            ' ExEnd:PrintHeaderUsingMhtFormatOptions
        End Sub
    End Class
End Namespace