Imports Aspose.Email.Mail

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email
    Class SetAlternateText
        Public Shared Sub Run()
            ' ExStart:SetAlternateText
            ' Declare message as MailMessage instance
            Dim message As New MailMessage()

            ' Creates AlternateView to view an email message using the content specified in the //string
            Dim alternate As AlternateView = AlternateView.CreateAlternateViewFromString("Alternate Text")

            ' Adding alternate text
            message.AlternateViews.Add(alternate)
            ' ExEnd:SetAlternateText
        End Sub
    End Class
End Namespace
