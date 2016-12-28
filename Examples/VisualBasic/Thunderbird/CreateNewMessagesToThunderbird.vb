Imports System.IO
Imports Aspose.Email.Formats.Mbox
Imports Aspose.Email.Mail

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email.Thunderbird
    Class CreateNewMessagesToThunderbird
        Public Shared Sub Run()
            ' ExStart:CreateNewMessagesToThunderbird
            ' The path to the File directory.
            Dim dataDir As String = RunExamples.Thunderbird()

            Try
                ' Open the storage file with FileStream
                Dim stream As New FileStream(dataDir + "Please add your Thunderbird file name here", FileMode.Open, FileAccess.Read)

                ' Initialize MboxStorageWriter and pass the above stream to it
                Dim writer As New MboxrdStorageWriter(stream, False)
                ' Prepare a new message using the MailMessage class
                Dim message As New MailMessage("from@domain.com", "to@domain.com", "added 1 from Aspose.Email", "added from Aspose.Email")
                ' Add this message to storage
                writer.WriteMessage(message)
                ' Close all related streams
                writer.Dispose()
                ' ExEnd:CreateNewMessagesToThunderbird
                stream.Close()
            Catch ex As Exception
                Console.WriteLine(ex.Message + vbLf & "Please add Thunderbird file name to the FileStream")
            End Try
        End Sub
    End Class
End Namespace