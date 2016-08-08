Imports System.IO
Imports Aspose.Email.Formats.Mbox
Imports Aspose.Email.Mail

Namespace Aspose.Email.Examples.VisualBasic.Email.Thunderbird
    Public Class GetMarkerInformation
        Public Shared Sub Run()
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
            'ExEnd: GetMarkerInformation
        End Sub
    End Class
End Namespace