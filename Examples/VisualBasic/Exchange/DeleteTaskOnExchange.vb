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
	Class DeleteTaskOnExchange
		Public Shared Sub Run()
			' ExStart:DeleteTaskOnExchange
			' Create instance of ExchangeClient class by giving credentials
			Dim client As IEWSClient = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain")

			' Get all tasks info collection from exchange
			Dim tasks As ExchangeMessageInfoCollection = client.ListMessages(client.MailboxInfo.TasksUri)

			' Parse all the tasks info in the list
			For Each info As ExchangeMessageInfo In tasks
				' Fetch task from exchange using current task info
				Dim task As ExchangeTask = client.FetchTask(info.UniqueUri)

				' Check if the current task fulfills the search criteria
				If task.Subject.Equals("test") Then
					' Delete task from exchange
					client.DeleteTask(task.UniqueUri, DeleteTaskOptions.DeletePermanently)
				End If
			Next
			' ExEnd:DeleteTaskOnExchange
		End Sub
	End Class
End Namespace