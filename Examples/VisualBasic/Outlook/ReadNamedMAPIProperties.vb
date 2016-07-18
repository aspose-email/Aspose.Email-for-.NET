Imports Aspose.Email.Outlook

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class ReadNamedMAPIProperties
        Public Shared Sub Run()

            'ExStart:ReadNamedMAPIProperties
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()

            ' Load from file
            Dim message As MapiMessage = MapiMessage.FromFile(dataDir & Convert.ToString("message.msg"))

            ' Get all named MAPI properties
            Dim properties As MapiPropertyCollection = message.NamedProperties
            ' Read all properties in foreach loop
            For Each mapiNamedProp As MapiNamedProperty In properties.Values
                ' Read any specific property
                Select Case mapiNamedProp.NameId
                    Case "TEST"
                        Console.WriteLine("{0} = {1}", mapiNamedProp.NameId, mapiNamedProp.GetString())
                        Exit Select
                    Case "MYPROP"
                        Console.WriteLine("{0} = {1}", mapiNamedProp.NameId, mapiNamedProp.GetString())
                        Exit Select
                    Case Else
                        Exit Select
                End Select
            Next
            'ExEnd:ReadNamedMAPIProperties
        End Sub
    End Class
End Namespace