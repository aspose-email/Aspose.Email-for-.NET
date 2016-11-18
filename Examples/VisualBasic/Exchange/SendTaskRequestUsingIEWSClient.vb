Imports System.Net
Imports Aspose.Email.Exchange
Imports Aspose.Email.Mail

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email.Exchange
    Class SendTaskRequestUsingIEWSClient
        Public Shared Sub Run()
            Try
                ' ExStart:SendTaskRequestUsingIEWSClient
                Dim dataDir As String = RunExamples.GetDataDir_Exchange()
                ' Create instance of ExchangeClient class by giving credentials
                Dim client As IEWSClient = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain")

                Dim options As New MsgLoadOptions()
                options.PreserveTnefAttachments = True

                ' load task from .msg file
                Dim eml As MailMessage = MailMessage.Load(dataDir & Convert.ToString("task.msg"), options)
                eml.From = "firstname.lastname@domain.com"
                eml.[To].Clear()
                eml.[To].Add(New MailAddress("firstname.lastname@domain.com"))
                ' ExEnd:SendTaskRequestUsingIEWSClient
                client.Send(eml)
            Catch ex As Exception
                Console.Write(ex.Message)
            End Try
        End Sub
    End Class
End Namespace