Imports System.IO
Imports Aspose.Email.Outlook

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class CreateAndSaveOutlookContact
        Public Shared Sub Run()
            ' ExStart:CreateAndSaveOutlookContact
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()

            Dim contact As New MapiContact()
            contact.NameInfo = New MapiContactNamePropertySet("Bertha", "A.", "Buell")
            contact.ProfessionalInfo = New MapiContactProfessionalPropertySet("Awthentikz", "Social work assistant")
            contact.PersonalInfo.PersonalHomePage = "B2BTies.com"
            contact.PhysicalAddresses.WorkAddress.Address = "Im Astenfeld 59 8580 EDELSCHROTT"
            contact.ElectronicAddresses.Email1 = New MapiContactElectronicAddress("Experwas", "SMTP", "BerthaABuell@armyspy.com")
            contact.Telephones = New MapiContactTelephonePropertySet("06605045265")
            contact.PersonalInfo.Children = New String() {"child1", "child2", "child3"}
            contact.Categories = New String() {"category1", "category2", "category3"}
            contact.Mileage = "Some test mileage"
            contact.Billing = "Test billing information"
            contact.OtherFields.Journal = True
            contact.OtherFields.[Private] = True
            contact.OtherFields.ReminderTime = New DateTime(2014, 1, 1, 0, 0, 55)
            contact.OtherFields.ReminderTopic = "Test topic"
            contact.OtherFields.UserField1 = "ContactUserField1"
            contact.OtherFields.UserField2 = "ContactUserField2"
            contact.OtherFields.UserField3 = "ContactUserField3"
            contact.OtherFields.UserField4 = "ContactUserField4"

            ' Add a photo
            Using fs As FileStream = File.OpenRead(dataDir & Convert.ToString("Desert.jpg"))
                Dim buffer As Byte() = New Byte(fs.Length - 1) {}
                fs.Read(buffer, 0, buffer.Length)
                contact.Photo = New MapiContactPhoto(buffer, MapiContactPhotoImageFormat.Jpeg)
            End Using
            ' Save the Contact in MSG format
            contact.Save(dataDir & Convert.ToString("MapiContact_out.msg"), ContactSaveFormat.Msg)

            ' Save the Contact in VCF format
            contact.Save(dataDir & Convert.ToString("MapiContact_out.vcf"), ContactSaveFormat.VCard)
            ' ExEnd:CreateAndSaveOutlookContact
        End Sub
    End Class
End Namespace