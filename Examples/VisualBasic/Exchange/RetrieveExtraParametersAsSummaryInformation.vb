Imports System.Threading
Imports Aspose.Email.Imap
Imports Aspose.Email.Mail

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'


Namespace Aspose.Email.Examples.CSharp.Email.IMAP
    Class RetrieveExtraParametersAsSummaryInformation
        Public Shared Sub Run()
            Try
                ' ExStart:RetrieveExtraParametersAsSummaryInformation
                Using client As New ImapClient("host.domain.com", "username", "password")
                    Dim message As New MailMessage("from@domain.com", "to@doman.com", "EMAILNET-38466 - " + Guid.NewGuid().ToString(), "EMAILNET-38466 Add extra parameters for UID FETCH command")

                    ' append the message to the server
                    Dim uid As String = client.AppendMessage(message)

                    ' wait for the message to be appended
                    Thread.Sleep(5000)

                    ' Define properties to be fetched from server along with the message
                    Dim messageExtraFields As String() = New String() {"X-GM-MSGID", "X-GM-THRID"}

                    ' retreive the message summary information using it's UID and sequence number
                    Dim messageInfoUID As ImapMessageInfo = client.ListMessage(uid, messageExtraFields)
                    Dim messageInfoSeqNum As ImapMessageInfo = client.ListMessage(1, messageExtraFields)

                    ' List messages in general from the server based on the defined properties
                    Dim messageInfoCol As ImapMessageInfoCollection = client.ListMessages(messageExtraFields)
                    Dim messageInfoFromList As ImapMessageInfo = messageInfoCol(0)

                    ' verify that the parameters are fetched in the summary information

                    For Each paramName As String In messageExtraFields
                        Console.WriteLine(messageInfoFromList.ExtraParameters.ContainsKey(paramName))
                        Console.WriteLine(messageInfoUID.ExtraParameters.ContainsKey(paramName))
                        Console.WriteLine(messageInfoSeqNum.ExtraParameters.ContainsKey(paramName))
                    Next
                    ' ExEnd:RetrieveExtraParametersAsSummaryInformation
                End Using
            Catch ex As Exception
                Console.Write(ex.Message)
                Throw
            End Try
        End Sub
    End Class
End Namespace