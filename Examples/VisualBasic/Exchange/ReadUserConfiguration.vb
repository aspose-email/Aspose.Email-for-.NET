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
	Class ReadUserConfiguration
		Public Shared Sub Run()
			Const  mailboxUri As String = "https://exchnage/ews/exchange.asmx"
			Const  domain As String = ""
			Const  username As String = "username@ASE305.onmicrosoft.com"
			Const  password As String = "password"
			Dim credentials As New NetworkCredential(username, password, domain)
			Dim client As IEWSClient = EWSClient.GetEWSClient(mailboxUri, credentials)
			Console.WriteLine("Connected to Exchange 2010")

			' ExStart:ReadUserConfiguration
			' Get the User Configuration for Inbox folder
			Dim userConfigName As New UserConfigurationName("inbox.config", client.MailboxInfo.InboxUri)
			Dim userConfig As UserConfiguration = client.GetUserConfiguration(userConfigName)

			Console.WriteLine("Configuration Id: " + userConfig.Id)
			Console.WriteLine("Configuration Name: " + userConfig.UserConfigurationName.Name)
			Console.WriteLine("Key value pairs:")
			For Each key As String In userConfig.Dictionary.Keys
				Console.WriteLine((key & Convert.ToString(": ")) + userConfig.Dictionary(key).ToString())
			Next
			' ExEnd:ReadUserConfiguration
		End Sub
	End Class
End Namespace