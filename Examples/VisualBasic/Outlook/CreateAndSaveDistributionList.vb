Imports Aspose.Email.Outlook

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class CreateAndSaveDistributionList
        Public Shared Sub Run()
            ' ExStart:CreateAndSaveDistributionList
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()

            ' Create MapiDistributionListMemberCollection object and add MapiDistributionListMembers
            Dim oneOffmembers As New MapiDistributionListMemberCollection()
            oneOffmembers.Add(New MapiDistributionListMember("John R. Patrick", "JohnRPatrick@armyspy.com"))
            oneOffmembers.Add(New MapiDistributionListMember("Tilly Bates", "TillyBates@armyspy.com"))

            Dim dlist As New MapiDistributionList("Simple list", oneOffmembers)
            ' Set MapiDistributionList properties
            dlist.Body = "Test body"
            dlist.Subject = "Test subject"
            dlist.Mileage = "Test mileage"
            dlist.Billing = "Test billing"

            dlist.Save(dataDir & Convert.ToString("distlist_out.msg"))
            ' ExEnd:CreateAndSaveDistributionList
        End Sub
    End Class
End Namespace
