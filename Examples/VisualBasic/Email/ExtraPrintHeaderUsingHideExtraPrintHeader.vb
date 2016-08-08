Imports System.IO
Imports System.Security.Cryptography.X509Certificates
Imports Aspose.Email.Mail

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email
	Class ExtraPrintHeaderUsingHideExtraPrintHeader
		Public Shared Sub Run()
			' ExStart:ExtraPrintHeaderUsingHideExtraPrintHeader
			' The path to the File directory.
			Dim dataDir As String = RunExamples.GetDataDir_Email()

			Dim mhtFileName As String = dataDir & Convert.ToString("Message.mhtml")
			Dim message As MailMessage = MailMessage.Load(dataDir & Convert.ToString("Message.eml"))
			Dim encodedPageHeader As String = "<div><div class=3D'page=Header'>&quotPanditharatne, Mithra&quot &ltmithra=2Epanditharatne@cibc==2Ecom&gt<hr/></div>"

			Dim mailFormatter As New MhtMessageFormatter()
			Dim options As MhtFormatOptions = MhtFormatOptions.WriteCompleteEmailAddress Or MhtFormatOptions.WriteHeader
			mailFormatter.Format(message)

			message.Save(mhtFileName, Aspose.Email.Mail.SaveOptions.DefaultMhtml)

			If File.ReadAllText(mhtFileName).Contains(encodedPageHeader) Then
				Console.WriteLine("True")
			Else
				Console.WriteLine("False")
			End If

			'Assert.True(File.ReadAllText(mhtFileName).Contains(encodedPageHeader))
			options = options Or MhtFormatOptions.HideExtraPrintHeader
			mailFormatter.Format(message)
			message.Save(mhtFileName, Aspose.Email.Mail.SaveOptions.DefaultMhtml)
			If File.ReadAllText(mhtFileName).Contains(encodedPageHeader) Then
				Console.WriteLine("True")
			Else
				Console.WriteLine("False")
			End If
			' ExEnd:ExtraPrintHeaderUsingHideExtraPrintHeader
		End Sub
	End Class
End Namespace