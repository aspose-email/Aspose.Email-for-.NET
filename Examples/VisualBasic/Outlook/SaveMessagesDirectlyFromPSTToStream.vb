Imports System.IO
Imports Aspose.Email.Outlook.Pst

' This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET 
'   API reference when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq 
'   for more information. If you do not wish to use NuGet, you can manually download 
'   Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'   install it and then add its reference to this project. For any issues, questions or suggestions 
'   please feel free to contact us using http://www.aspose.com/community/forums/default.aspx            
'


Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class SaveMessagesDirectlyFromPSTToStream
        Public Shared Sub Run()
            ' ExStart:SaveMessagesDirectlyFromPSTToStream
            ' The path to the file directory.
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()

            ' Load the Outlook file
            Dim path As String = dataDir & Convert.ToString("PersonalStorage.pst")

            ' Save message to file
            Using personalStorage__1 As PersonalStorage = PersonalStorage.FromFile(path)
                Dim inbox As FolderInfo = personalStorage__1.RootFolder.GetSubFolder("Inbox")
                For Each messageInfo As MessageInfo In inbox.EnumerateMessages()
                    Using fs As FileStream = File.OpenWrite(messageInfo.Subject + ".msg")
                        personalStorage__1.SaveMessageToStream(messageInfo.EntryIdString, fs)
                    End Using
                Next
            End Using
            ' ExEnd:SaveMessagesDirectlyFromPSTToStream
        End Sub
    End Class
End Namespace
