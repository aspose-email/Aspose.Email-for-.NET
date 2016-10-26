Imports Aspose.Email.Mail

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class Exposeproperties
        Public Shared Sub Run()
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()

            ' Load mail message
            Dim msg As MailMessage = MailMessage.Load(dataDir & Convert.ToString("Message.eml"), New EmlLoadOptions())

            ' Subject
            If msg.Subject IsNot Nothing Then
                Console.WriteLine(msg.Subject)
            Else
                Console.WriteLine("no subject")
            End If

            ' From
            If msg.From IsNot Nothing Then
                Console.WriteLine(msg.From)
            Else
                Console.WriteLine("No sender")
            End If

            ' To
            If msg.[To] IsNot Nothing Then
                Console.WriteLine(msg.[To])
            Else
                Console.WriteLine("No one in To")
            End If

            ' Cc
            If msg.CC IsNot Nothing Then
                Console.WriteLine(msg.CC)
            Else
                Console.WriteLine("No one in cc")
            End If
        End Sub
    End Class
End Namespace