Imports System.Text
Imports Aspose.Email
Imports Aspose.Email.Outlook
Imports Aspose.Email.Outlook.Pst

' This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET 
'   API reference when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq 
'   for more information. If you do not wish to use NuGet, you can manually download 
'   Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'   install it and then add its reference to this project. For any issues, questions or suggestions 
'   please feel free to contact us using http://www.aspose.com/community/forums/default.aspx            
'

Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class UpdatePSTCustomProperites
        ' ExStart:UpdatePSTCustomProperites
        Public Shared Sub Run()
            ' Load the Outlook file
            Dim dataDir As String = RunExamples.GetDataDir_Outlook() + "Outlook.pst"
            Using personalStorage__1 As PersonalStorage = PersonalStorage.FromFile(dataDir)
                Dim testFolder As FolderInfo = personalStorage__1.RootFolder.GetSubFolder("Inbox")

                ' Create the collection of message properties for adding or updating
                Dim newProperties As New MapiPropertyCollection()

                ' Normal,  Custom and PidLidLogFlags named  property
                Dim [property] As New MapiProperty(MapiPropertyTag.PR_ORG_EMAIL_ADDR_W, Encoding.Unicode.GetBytes("test_address@org.com"))
                Dim namedProperty1 As MapiProperty = New MapiNamedProperty(GenerateNamedPropertyTag(0, MapiPropertyType.PT_LONG), "ITEM_ID", Guid.NewGuid(), BitConverter.GetBytes(123))
                Dim namedProperty2 As MapiProperty = New MapiNamedProperty(GenerateNamedPropertyTag(1, MapiPropertyType.PT_LONG), &H870C, New Guid("0006200A-0000-0000-C000-000000000046"), BitConverter.GetBytes(0))
                newProperties.Add(namedProperty1.Tag, namedProperty1)
                newProperties.Add(namedProperty2.Tag, namedProperty2)
                newProperties.Add([property].Tag, [property])
                testFolder.ChangeMessages(testFolder.EnumerateMessagesEntryId(), newProperties)
            End Using
        End Sub

        Private Shared Function GenerateNamedPropertyTag(index As Long, dataType As MapiPropertyType) As Long
            Return (CLng(CLng(dataType)) Or (&H8000 Or index) << 16) And &HFFFFFFFFUI
        End Function
        ' ExEnd:UpdatePSTCustomProperites
    End Class
End Namespace
