Imports Aspose.Email.Mail

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email
    Class SetHTMLBody
        Public Shared Sub Run()
            ' ExStart:SetHTMLBody
            ' Declare message as MailMessage instance
            Dim message As New MailMessage()

            ' Specify HtmlBody
            message.HtmlBody = "<html><body>This is the HTML body</body></html>"
            ' ExEnd:SetHTMLBody
        End Sub
    End Class
End Namespace