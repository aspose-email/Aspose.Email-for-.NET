Imports Aspose.Email.Outlook

Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook

    Public Class SaveMsgAsTemplate
        Private Shared Sub Run()
            'ExStart: SaveMsgAsTemplate
            '<summary>
            'This example shows how to save an Outlook MSG file to Outlook Template  using the MapiMessage API
            'Available from Aspose.Email for .NET 6.4.0 onwards
            'MapiMessage.SaveAsTemplate(Stream stream) - Saves to the specified stream as Outlook File Template(OFT format).
            'MapiMessage.SaveAsTemplate(string fileName) - Saves to the specified file as Outlook File Template(OFT format).
            '</summary>
            Using mapi As New MapiMessage("test@from.to", "test@to.to", "template subject", "Template body")
                Dim oftMapiFileName As String = "mapiToOft.msg"
                mapi.SaveAsTemplate(oftMapiFileName)
            End Using
            'ExEnd: SaveMsgAsTemplate
        End Sub

    End Class
End Namespace
