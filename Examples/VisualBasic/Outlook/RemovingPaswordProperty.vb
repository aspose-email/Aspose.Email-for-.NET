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
    Class RemovingPaswordProperty
        Public Shared Sub Run()
            ' The path to the documents directory.
            ' ExStart:RemovingPaswordProperty
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()
            Dim personalStorage__1 As PersonalStorage = PersonalStorage.FromFile(dataDir & Convert.ToString("PersonalStorage1.pst"))
            If personalStorage__1.Store.Properties.Contains(MapiPropertyTag.PR_PST_PASSWORD) Then
                Dim [property] As New MapiProperty(MapiPropertyTag.PR_PST_PASSWORD, BitConverter.GetBytes(CLng(0)))
                personalStorage__1.Store.SetProperty([property])
            End If
            ' ExEnd:RemovingPaswordProperty
        End Sub
    End Class
End Namespace
