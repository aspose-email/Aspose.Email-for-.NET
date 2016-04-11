Imports System.IO
Imports Aspose.Email.Formats.Mbox
Imports Aspose.Email.Mail

Public Class GetMarkerInformation
    Private Shared Sub Run()
        'ExStart: GetMarkerInformation
        '<summary>
        'This exmaple shows how to get marker information while reading or writing a message to Mbox file.
        'Introduced in Aspose.Email for .NET 6.4.0
        '</summary>
        Using stream As New FileStream("inbox.dat", FileMode.Open, FileAccess.Read)
            Using reader As New MboxrdStorageReader(stream, False)
                Dim fromMarker As String = Nothing
                Dim msg As MailMessage = reader.ReadNextMessage(fromMarker)
                'While (InlineAssignHelper(msg, reader.ReadNextMessage(fromMarker))) IsNot Nothing
                Do While Not msg Is Nothing

                    Console.WriteLine(fromMarker)

                    msg.Dispose()
                    msg = reader.ReadNextMessage(fromMarker)
                Loop
        End Using
        End Using
        Using writeStream As New FileStream("inbox.dat", FileMode.Create, FileAccess.Write)
            Using writer As New MboxrdStorageWriter(writeStream, False)
                Dim fromMarker As String = Nothing
                Dim msg As MailMessage = MailMessage.Load("input.eml")
                writer.WriteMessage(msg, fromMarker)

                Console.WriteLine(fromMarker)
            End Using
        End Using
        'ExEnd: GetMarkerInformaiton
    End Sub
End Class
