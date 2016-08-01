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
	Class AddMembersToPrivateDistributionList
		Public Shared Sub Run()
			' ExStart:AddMembersToPrivateDistributionList
			Dim client As IEWSClient = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain")
			Dim distributionLists As ExchangeDistributionList() = client.ListDistributionLists()
			Dim newMembers As New MailAddressCollection()
			newMembers.Add("address4@host.com")
			newMembers.Add("address5@host.com")
			client.AddToDistributionList(distributionLists(0), newMembers)
			' ExEnd:AddMembersToPrivateDistributionList
		End Sub
	End Class
End Namespace