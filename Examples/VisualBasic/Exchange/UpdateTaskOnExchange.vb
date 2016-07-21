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
	Class UpdateTaskOnExchange
		Public Shared Sub Run()
			' ExStart:UpdateTaskOnExchange
			' Create and initialize credentials
			Dim credentials = New NetworkCredential("username", "12345")

			' Create instance of ExchangeClient class by giving credentials
			Dim client As IEWSClient = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain")

			' Get all tasks info collection from exchange
			Dim tasks As ExchangeMessageInfoCollection = client.ListMessages(client.MailboxInfo.TasksUri)

			' Parse all the tasks info in the list
			For Each info As ExchangeMessageInfo In tasks
				' Fetch task from exchange using current task info
				Dim task As ExchangeTask = client.FetchTask(info.UniqueUri)

				' Update the task status to NotStarted
				task.Status = ExchangeTaskStatus.NotStarted

				' Set the task due date
				task.DueDate = New DateTime(2013, 2, 26)

				' Set task priority
				task.Priority = MailPriority.Low

				' Update task on exchange
				client.UpdateTask(task)
			Next
			' ExEnd:UpdateTaskOnExchange
		End Sub
	End Class
End Namespace