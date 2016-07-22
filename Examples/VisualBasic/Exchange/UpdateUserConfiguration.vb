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
	Class UpdateUserConfiguration
		Public Shared Sub Run()
			Const  mailboxUri As String = "https://exchnage/ews/exchange.asmx"
			Const  domain As String = ""
			Const  username As String = "username@ASE305.onmicrosoft.com"
			Const  password As String = "password"
			' ExStart:UpdatUserConfiguration
			Dim credentials As New NetworkCredential(username, password, domain)
			Dim client As IEWSClient = EWSClient.GetEWSClient(mailboxUri, credentials)
			Console.WriteLine("Connected to Exchange 2010")
			' Create the User Configuration for Inbox folder
			Dim userConfigName As New UserConfigurationName("inbox.config", client.MailboxInfo.InboxUri)
			Dim userConfig As New UserConfiguration(userConfigName)
			userConfig.Dictionary.Add("key1", "value1")
			userConfig.Dictionary.Add("key2", "value2")
			userConfig.Dictionary.Add("key3", "value3")
			client.CreateUserConfiguration(userConfig)
			' ExEnd:UpdatUserConfiguration
		End Sub
	End Class
End Namespace