Imports Aspose.Email.Outlook

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class ReadiNamedMAPIPropertyFromAttachment
        Public Shared Sub Run()
            ' ExStart:ReadiNamedMAPIPropertyFromAttachment
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()

            ' Load from file
            Dim msg As MapiMessage = MapiMessage.FromFile(dataDir & Convert.ToString("message.msg"))

            Dim subject As String

            ' Access the MapiPropertyTag.PR_SUBJECT property
            Dim prop As MapiProperty = msg.Properties(MapiPropertyTag.PR_SUBJECT)

            ' If the property is not found, check the MapiPropertyTag.PR_SUBJECT_W (which is a // Unicode peer of the MapiPropertyTag.PR_SUBJECT)
            If prop Is Nothing Then
                prop = msg.Properties(MapiPropertyTag.PR_SUBJECT_W)
            End If

            ' Cannot found
            If prop Is Nothing Then
                Console.WriteLine("No property found!")
                Return
            End If

            ' Get the property data as string
            subject = prop.GetString()

            Console.WriteLine(Convert.ToString("Subject:") & subject)
            ' Read internet code page property
            prop = msg.Properties(MapiPropertyTag.PR_INTERNET_CPID)
            If prop IsNot Nothing Then
                Console.WriteLine("CodePage:" + prop.GetLong().ToString())
            End If
            ' ExEnd:ReadiNamedMAPIPropertyFromAttachment
        End Sub
    End Class
End Namespace
