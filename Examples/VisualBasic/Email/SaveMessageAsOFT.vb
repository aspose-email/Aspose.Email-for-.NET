Imports Aspose.Email.Mail

Namespace Aspose.Email.Examples.VisualBasic.Email
    Public Class SaveMessageAsOFT
        Private Shared Sub Run()
            'ExStart: SaveMessageAsOFT
            ' <summary>
            ' This exmpale shows how to save an email as Outlook Template (.OFT) file using the MailMesasge API
            ' MsgSaveOptions.SaveAsTemplate - Set to true, if need to be saved as Outlook File Template(OFT format).
            ' Available from Aspose.Email for .NET 6.4.0 onwards
            ' </summary>

            Using eml As New MailMessage("test@from.to", "test@to.to", "template subject", "Template body")
                Dim oftEmlFileName As String = "EmlAsOft.oft"
                Dim options As MsgSaveOptions = SaveOptions.DefaultMsgUnicode
                options.SaveAsTemplate = True
                eml.Save(oftEmlFileName, options)
            End Using

            'ExEnd: SaveMessageAsOFT
        End Sub
    End Class
End Namespace
