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
    Class CreateNewMapiContactAndAddToContactsSubfolder
        Public Shared Sub Run()
            ' The path to the documents directory.
            ' ExStart:CreateNewMapiContactAndAddToContactsSubfolder
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()

            ' Create three Contacts 
            Dim contact1 As New MapiContact("Sebastian Wright", "SebastianWright@dayrep.com")
            Dim contact2 As New MapiContact("Wichert Kroos", "WichertKroos@teleworm.us", "Grade A Investment")
            Dim contact3 As New MapiContact("Christoffer van de Meeberg", "ChristoffervandeMeeberg@teleworm.us", "Krauses Sofa Factory", "046-630-4614046-630-4614")

            ' Contact #4
            Dim contact4 As New MapiContact()
            contact4.NameInfo = New MapiContactNamePropertySet("Margaret", "J.", "Tolle")
            contact4.PersonalInfo.Gender = MapiContactGender.Female
            contact4.ProfessionalInfo = New MapiContactProfessionalPropertySet("Adaptaz", "Recording engineer")
            contact4.PhysicalAddresses.WorkAddress.Address = "4 Darwinia Loop EIGHTY MILE BEACH WA 6725"
            contact4.ElectronicAddresses.Email1 = New MapiContactElectronicAddress("Hisen1988", "SMTP", "MargaretJTolle@dayrep.com")
            contact4.Telephones.BusinessTelephoneNumber = "(08)9080-1183"
            contact4.Telephones.MobileTelephoneNumber = "(925)599-3355(925)599-3355"

            ' Contact #5
            Dim contact5 As New MapiContact()
            contact5.NameInfo = New MapiContactNamePropertySet("Matthew", "R.", "Wilcox")
            contact5.PersonalInfo.Gender = MapiContactGender.Male
            contact5.ProfessionalInfo = New MapiContactProfessionalPropertySet("Briazz", "Psychiatric aide")
            contact5.PhysicalAddresses.WorkAddress.Address = "Horner Strasse 12 4421 SAASS"
            contact5.Telephones.BusinessTelephoneNumber = "0650 675 73 300650 675 73 30"
            contact5.Telephones.HomeTelephoneNumber = "(661)387-5382(661)387-5382"

            ' Contact #6
            Dim contact6 As New MapiContact()
            contact6.NameInfo = New MapiContactNamePropertySet("Bertha", "A.", "Buell")
            contact6.ProfessionalInfo = New MapiContactProfessionalPropertySet("Awthentikz", "Social work assistant")
            contact6.PersonalInfo.PersonalHomePage = "B2BTies.com"
            contact6.PhysicalAddresses.WorkAddress.Address = "Im Astenfeld 59 8580 EDELSCHROTT"
            contact6.ElectronicAddresses.Email1 = New MapiContactElectronicAddress("Experwas", "SMTP", "BerthaABuell@armyspy.com")
            contact6.Telephones = New MapiContactTelephonePropertySet("06605045265")

            Dim checkfile As String = dataDir + "SampleContacts_out.pst"

            If File.Exists(checkfile) Then
                File.Delete(checkfile)
            Else
            End If



            Using personalStorage__1 As PersonalStorage = PersonalStorage.Create(dataDir & Convert.ToString("SampleContacts_out.pst"), FileFormatVersion.Unicode)
                Dim contactFolder As FolderInfo = personalStorage__1.CreatePredefinedFolder("Contacts", StandardIpmFolder.Contacts)
                contactFolder.AddMapiMessageItem(contact1)
                contactFolder.AddMapiMessageItem(contact2)
                contactFolder.AddMapiMessageItem(contact3)
                contactFolder.AddMapiMessageItem(contact4)
                contactFolder.AddMapiMessageItem(contact5)
                contactFolder.AddMapiMessageItem(contact6)
            End Using
            ' ExEnd:CreateNewMapiContactAndAddToContactsSubfolder
        End Sub
    End Class
End Namespace
