Imports Aspose.Email.Outlook

Namespace Aspose.Email.Examples.VisualBasic.Email.POP3
	Class RecipientInformation
		Public Shared Sub Run()
			' ExStart:RecipientInformation
			' The path to the File directory.
			Dim dataDir As String = RunExamples.GetDataDir_POP3()
            Dim dstEmail As String = dataDir & Convert.ToString("Message.msg")

			' Load message file and Enumerate the recipients
			Dim message As MapiMessage = MapiMessage.FromFile(dstEmail)
			For Each recip As MapiRecipient In message.Recipients
				Select Case recip.RecipientType
					' What's the type?
					Case MapiRecipientType.MAPI_TO
						Console.WriteLine("RecipientType:TO")
						Exit Select
					Case MapiRecipientType.MAPI_CC
						Console.WriteLine("RecipientType:CC")
						Exit Select
					Case MapiRecipientType.MAPI_BCC
						Console.WriteLine("RecipientType:BCC")
						Exit Select
				End Select
				' Get email address, display name and  address type
				Console.WriteLine("Email Address: " + recip.EmailAddress)
				Console.WriteLine("DisplayName: " + recip.DisplayName)
				Console.WriteLine("AddressType: " + recip.AddressType)
			Next
			' ExEnd:RecipientInformation
			Console.WriteLine(Convert.ToString(Environment.NewLine + "Displayed recipient information from MSG file ") & dstEmail)
		End Sub
	End Class
End Namespace