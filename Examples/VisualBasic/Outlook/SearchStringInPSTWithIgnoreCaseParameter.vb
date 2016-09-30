Imports System.IO
Imports Aspose.Email.Outlook
Imports Aspose.Email.Outlook.Pst
Imports Aspose.Email.Mail

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class SearchStringInPSTWithIgnoreCaseParameter
        Public Shared Sub Run()
            ' The path to the File directory.
            ' ExStart:SearchStringInPSTWithIgnoreCaseParameter
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()
            Dim checkfile As String = dataDir + "SearchStringInPSTWithIgnoreCaseParameter_out.pst"

            If File.Exists(checkfile) Then
                File.Delete(checkfile)
            Else
            End If

            Using personalStorage1 As PersonalStorage = PersonalStorage.Create(dataDir & Convert.ToString("SearchStringInPSTWithIgnoreCaseParameter_out.pst"), FileFormatVersion.Unicode)
                Dim folderInfo As FolderInfo = personalStorage1.CreatePredefinedFolder("Inbox", StandardIpmFolder.Inbox)

                folderInfo.AddMessage(MapiMessage.FromMailMessage(MailMessage.Load(dataDir + "Message.eml")))

                Dim builder As New PersonalStorageQueryBuilder()
                ' IgnoreCase is True
                builder.From.Contains("automated", True)

                Dim query As MailQuery = builder.GetQuery()
                Dim coll As MessageInfoCollection = folderInfo.GetContents(query)
                Console.WriteLine(coll.Count.ToString())
            End Using
            ' ExEnd:SearchStringInPSTWithIgnoreCaseParameter
        End Sub
    End Class
End Namespace
