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
    Class ReadMessagesFromThunderbird
        Public Shared Sub Run()
            ' ExStart:ReadMessagesFromThunderbird

            ' The path to the File directory.
            Dim dataDir As String = RunExamples.Thunderbird()

            ' Open the storage file with FileStream
            Dim stream As New FileStream(dataDir & Convert.ToString("Outlook.pst"), FileMode.Open, FileAccess.Read)
            ' Create an instance of the MboxrdStorageReader class and pass the stream
            Dim reader As New MboxrdStorageReader(stream, False)
            ' Start reading messages
            Dim message As MailMessage = reader.ReadNextMessage()

            ' Read all messages in a loop
            While message IsNot Nothing
                ' Manipulate message - show contents
                Console.WriteLine("Subject: " + message.Subject)
                ' Save this message in EML or MSG format
                message.Save(message.Subject + ".eml", SaveOptions.DefaultEml)
                message.Save(message.Subject + ".msg", SaveOptions.DefaultMsgUnicode)

                ' Get the next message
                message = reader.ReadNextMessage()
            End While
            ' Close the streams
            reader.Dispose()
            stream.Close()
            ' ExEnd:ReadMessagesFromThunderbird
        End Sub
    End Class
End Namespace