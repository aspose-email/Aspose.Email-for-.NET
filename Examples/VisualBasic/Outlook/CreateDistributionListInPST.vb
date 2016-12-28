Imports System.IO
Imports Aspose.Email.Outlook
Imports Aspose.Email.Outlook.Pst

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class CreateDistributionListInPST
        Public Shared Sub Run()
            ' The path to the File directory.
            ' ExStart:CreateDistributionListInPST
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()

            Dim displayName1 As String = "Sebastian Wright"
            Dim email1 As String = "SebastianWright@dayrep.com"

            Dim displayName2 As String = "Wichert Kroos"
            Dim email2 As String = "WichertKroos@teleworm.us"

            Dim strEntryId1 As String
            Dim strEntryId2 As String

            Dim checkfile As String = dataDir + "CreateDistributionListInPST_out.pst"

            If File.Exists(checkfile) Then
                File.Delete(checkfile)
            Else
            End If

            ' Create distribution list from contacts
            Using personalStorage__1 As PersonalStorage = PersonalStorage.Create(dataDir & Convert.ToString("CreateDistributionListInPST_out.pst"), FileFormatVersion.Unicode)
                ' Add the contact folder to pst
                Dim contactFolder As FolderInfo = personalStorage__1.CreatePredefinedFolder("Contacts", StandardIpmFolder.Contacts)

                ' Create contacts
                strEntryId1 = contactFolder.AddMapiMessageItem(New MapiContact(displayName1, email1))
                strEntryId2 = contactFolder.AddMapiMessageItem(New MapiContact(displayName2, email2))

                ' Create distribution list on the base of the created contacts
                Dim member1 As New MapiDistributionListMember(displayName1, email1)
                member1.EntryIdType = MapiDistributionListEntryIdType.Contact
                member1.EntryId = Convert.FromBase64String(strEntryId1)

                Dim member2 As New MapiDistributionListMember(displayName2, email2)
                member2.EntryIdType = MapiDistributionListEntryIdType.Contact
                member2.EntryId = Convert.FromBase64String(strEntryId2)

                Dim members As New MapiDistributionListMemberCollection()
                members.Add(member1)
                members.Add(member2)

                Dim distributionList As New MapiDistributionList("Contact list", members)
                distributionList.Body = "Distribution List Body"
                distributionList.Subject = "Sample Distribution List using Aspose.Email"

                ' Add distribution list to PST
                contactFolder.AddMapiMessageItem(distributionList)
            End Using

            Dim checkfile1 As String = dataDir + "CreateDistributionListInPST_OneOffmembers_out.pst"
            If File.Exists(checkfile1) Then
                File.Delete(checkfile1)
            Else
            End If

            ' Create one-off distribution list members (for which no separate contacts were created)
            Using personalStorage__1 As PersonalStorage = PersonalStorage.Create(dataDir & Convert.ToString("CreateDistributionListInPST_OneOffmembers_out.pst"), FileFormatVersion.Unicode)
                ' Add the contact folder to pst
                Dim contactFolder As FolderInfo = personalStorage__1.CreatePredefinedFolder("Contacts", StandardIpmFolder.Contacts)

                Dim oneOffmembers As New MapiDistributionListMemberCollection()
                oneOffmembers.Add(New MapiDistributionListMember("John R. Patrick", "JohnRPatrick@armyspy.com"))
                oneOffmembers.Add(New MapiDistributionListMember("Tilly Bates", "TillyBates@armyspy.com"))

                Dim oneOffMembersList As New MapiDistributionList("Simple list", oneOffmembers)
                contactFolder.AddMapiMessageItem(oneOffMembersList)
            End Using
            ' ExEnd:CreateDistributionListInPST
        End Sub
    End Class
End Namespace
