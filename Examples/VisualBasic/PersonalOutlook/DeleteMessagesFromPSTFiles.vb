Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports Aspose.Email.Outlook.Pst
Imports Aspose.Email

' This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET 
'   API reference when the project is build. Please check https://docs.nuget.org/consume/nuget-faq 
'   for more information. If you do not wish to use NuGet, you can manually download 
'   Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'   install it and then add its reference to this project. For any issues, questions or suggestions 
'   please feel free to contact us using http://www.aspose.com/community/forums/default.aspx            
'


Namespace Aspose.Email.Examples.VisualBasic.Email.Personal.Outlook
    Class DeleteMessagesFromPSTFiles
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()

            ' Load the Outlook PST file
            Dim personalStorage1 As PersonalStorage = PersonalStorage.FromFile(dataDir & Convert.ToString("Test.pst"))

            ' Get the Sent items folder
            Dim folderInfo As FolderInfo = personalStorage1.GetPredefinedFolder(StandardIpmFolder.SentItems)

            Dim msgInfoColl As MessageInfoCollection = folderInfo.GetContents()
            For Each msgInfo As MessageInfo In msgInfoColl
                Console.WriteLine(msgInfo.Subject + ": " + msgInfo.EntryIdString)
                If msgInfo.Subject.Equals("some delete condition") = True Then
                    ' Delete this item
                    folderInfo.DeleteChildItem(msgInfo.EntryId)
                    Console.WriteLine("Deleted this message")
                End If
            Next
        End Sub
    End Class
End Namespace
