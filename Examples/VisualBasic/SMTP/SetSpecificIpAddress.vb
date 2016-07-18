Imports System.Net
Imports Aspose.Email.Mail
Imports Aspose.Email

'  This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference when the project is build. 
'  Please check https:docs.nuget.org/consume/nuget-faq for more information. If you do not wish to use NuGet, 
'  you can manually download Aspose.Email for .NET API from http:www.aspose.com/downloads, 
'  install it and then add its reference to this project. 
'  For any issues, questions or suggestions please feel free to contact us 
'  using http:www.aspose.com/community/forums/default.aspx

Namespace Aspose.Email.Examples.VisualBasic.Email.SMTP

    Public Class SetSpecificIpAddress
        Public Shared Sub Run()
            Using client As New SmtpClient("smtp.domain.com", 587, "username", "password", SecurityOptions.Auto)
                ' set callback
                AddHandler client.BindIPEndPoint, AddressOf BindIPEndPointCallback
                client.Noop()
            End Using

        End Sub
        Private Shared Function BindIPEndPointCallback(remoteEndPoint As IPEndPoint) As IPEndPoint
            Return New IPEndPoint(IPAddress.Any, 0)
        End Function
    End Class
End Namespace