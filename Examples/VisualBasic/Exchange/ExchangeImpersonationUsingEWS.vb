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
	Class ExchangeImpersonationUsingEWS
		Public Shared Sub Run()
			' ExStart:ExchangeImpersonationUsingEWS
			' Create instance's of EWSClient class by giving credentials
			Dim client1 As IEWSClient = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser1", "pwd", "domain")
			Dim client2 As IEWSClient = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser2", "pwd", "domain")
			If True Then
				Dim folder As String = "Drafts"
				Try
					For Each messageInfo As ExchangeMessageInfo In client1.ListMessages(folder)
						client1.DeleteMessage(messageInfo.UniqueUri)
					Next
					Dim subj1 As String = String.Format("NETWORKNET_33354 {0} {1}", "User", "User1")
					client1.AppendMessage(folder, New MailMessage("User1@exchange.conholdate.local", "To@aspsoe.com", subj1, ""))

					For Each messageInfo As ExchangeMessageInfo In client2.ListMessages(folder)
						client2.DeleteMessage(messageInfo.UniqueUri)
					Next
					Dim subj2 As String = String.Format("NETWORKNET_33354 {0} {1}", "User", "User2")
					client2.AppendMessage(folder, New MailMessage("User2@exchange.conholdate.local", "To@aspose.com", subj2, ""))

					Dim messInfoColl As ExchangeMessageInfoCollection = client1.ListMessages(folder)
					'Assert.AreEqual(1, messInfoColl.Count)
					'Assert.AreEqual(subj1, messInfoColl[0].Subject)

					client1.ImpersonateUser(ItemChoice.PrimarySmtpAddress, "User2@exchange.conholdate.local")
					Dim messInfoColl1 As ExchangeMessageInfoCollection = client1.ListMessages(folder)
					'Assert.AreEqual(1, messInfoColl1.Count)
					'Assert.AreEqual(subj2, messInfoColl1[0].Subject)

					client1.ResetImpersonation()
						'Assert.AreEqual(1, messInfoColl2.Count)
						'Assert.AreEqual(subj1, messInfoColl2[0].Subject)
					Dim messInfoColl2 As ExchangeMessageInfoCollection = client1.ListMessages(folder)
				Finally
					Try
						For Each messageInfo As ExchangeMessageInfo In client1.ListMessages(folder)
							client1.DeleteMessage(messageInfo.UniqueUri)
						Next
						For Each messageInfo As ExchangeMessageInfo In client2.ListMessages(folder)
							client2.DeleteMessage(messageInfo.UniqueUri)
						Next
					Catch
					End Try
				End Try
			End If
			' ExEnd:ExchangeImpersonationUsingEWS
		End Sub
	End Class
End Namespace