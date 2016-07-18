Imports Aspose.Email.Outlook
Imports Aspose.Email.Outlook.Pst

' This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET 
'   API reference when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq 
'   for more information. If you do not wish to use NuGet, you can manually download 
'   Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'   install it and then add its reference to this project. For any issues, questions or suggestions 
'   please feel free to contact us using http://www.aspose.com/community/forums/default.aspx            
'

Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class SaveCalendarItems
        Public Shared Sub Run()
            ' ExStart:SaveCalendarItems
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()

            ' Load the Outlook PST file
            Dim pst As PersonalStorage = PersonalStorage.FromFile(dataDir & Convert.ToString("Sub.pst"))
            ' Get the Calendar folder
            Dim folderInfo As FolderInfo = pst.RootFolder.GetSubFolder("Inbox")
            ' Loop through all the calendar items in this folder
            Dim messageInfoCollection As MessageInfoCollection = folderInfo.GetContents()
            For Each messageInfo As MessageInfo In messageInfoCollection
                ' Get the calendar information
                Dim calendar As MapiMessage = DirectCast(pst.ExtractMessage(messageInfo).ToMapiMessageItem(), MapiMessage)
                ' Display some contents on screen
                Console.WriteLine("Name: " + calendar.Subject)
                ' Save to disk in ICS format
                calendar.Save((dataDir & Convert.ToString("\Calendar\")) + calendar.Subject + "_out.ics")
            Next
            ' ExEnd:SaveCalendarItems
        End Sub
    End Class
End Namespace
