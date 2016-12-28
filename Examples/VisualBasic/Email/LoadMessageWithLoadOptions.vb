Imports System.Text
Imports Aspose.Email.Mail

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email
	Class LoadMessageWithLoadOptions
		Public Shared Sub Run()
			' ExStart:LoadMessageWithLoadOptions
			' The path to the File directory.
			Dim dataDir As String = RunExamples.GetDataDir_Email()

			' Load Eml, html, mhtml, msg and dat file 
			Dim mailMessage__1 As MailMessage = MailMessage.Load(dataDir & Convert.ToString("Message.eml"), New EmlLoadOptions())
			MailMessage.Load(dataDir & Convert.ToString("description.html"), New HtmlLoadOptions())
			MailMessage.Load(dataDir & Convert.ToString("Message.mhtml"), New MhtmlLoadOptions())
			MailMessage.Load(dataDir & Convert.ToString("Message.msg"), New MsgLoadOptions())

			' loading with custom options
            Dim emlLoadOptions As New EmlLoadOptions() With { _
                .PrefferedTextEncoding = Encoding.UTF8, _
                .PreserveTnefAttachments = True _
            }

			MailMessage.Load(dataDir & Convert.ToString("description.html"), emlLoadOptions)
            Dim htmlLoadOptions As New HtmlLoadOptions() With { _
                .PrefferedTextEncoding = Encoding.UTF8, _
                .ShouldAddPlainTextView = True, _
                .PathToResources = dataDir _
            }
			MailMessage.Load(dataDir & Convert.ToString("description.html"), emlLoadOptions)
			' ExEnd:LoadMessageWithLoadOptions
		End Sub
	End Class
End Namespace