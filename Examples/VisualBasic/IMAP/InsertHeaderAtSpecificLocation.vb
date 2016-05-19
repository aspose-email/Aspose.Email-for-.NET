'  This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference when the project is build. 
'  Please check https://docs.nuget.org/consume/nuget-faq for more information. If you do not wish to use NuGet, 
'  you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'  install it and then add its reference to this project. 
'  For any issues, questions or suggestions please feel free to contact us 
'  using http://www.aspose.com/community/forums/default.aspx


Imports Aspose.Email.Mail

Public Class InsertHeaderAtSpecificLocation
    Public Shared Sub Run()
        Dim dataDir As String = RunExamples.GetDataDir_IMAP()
        Dim loadFile As String = dataDir & Convert.ToString("MessageHeaders.eml")
        Dim eml As MailMessage = MailMessage.Load(loadFile)
        eml.Headers.Insert("secret-header", "mystery1")
        eml.Save(dataDir & Convert.ToString("Updated-MessageHeaders.eml"))
    End Sub
End Class
