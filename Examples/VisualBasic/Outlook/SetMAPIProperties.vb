Imports System.Text
Imports Aspose.Email.Outlook

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class SetMAPIProperties
        Public Shared Sub Run()
            'ExStart:SetMAPIProperties
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()

            ' Create a sample Message
            Dim mapiMsg As New MapiMessage("user1@gmail.com", "user2@gmail.com", "This is subject", "This is body")

            ' Set multiple properties
            mapiMsg.SetProperty(New MapiProperty(MapiPropertyTag.PR_SENDER_ADDRTYPE_W, Encoding.Unicode.GetBytes("EX")))
            Dim recipientTo As MapiRecipient = mapiMsg.Recipients(0)
            Dim propAddressType As New MapiProperty(MapiPropertyTag.PR_RECEIVED_BY_ADDRTYPE_W, Encoding.UTF8.GetBytes("MYFAX"))
            recipientTo.SetProperty(propAddressType)
            Dim faxAddress As String = "My Fax User@/FN=fax#/VN=voice#/CO=My Company/CI=Local"
            Dim propEmailAddress As New MapiProperty(MapiPropertyTag.PR_RECEIVED_BY_EMAIL_ADDRESS_W, Encoding.UTF8.GetBytes(faxAddress))
            recipientTo.SetProperty(propEmailAddress)
            mapiMsg.SetMessageFlags(MapiMessageFlags.MSGFLAG_UNSENT Or MapiMessageFlags.MSGFLAG_FROMME)
            mapiMsg.SetProperty(New MapiProperty(MapiPropertyTag.PR_RTF_IN_SYNC, BitConverter.GetBytes(CLng(1))))

            ' Set DateTime property
            Dim modificationTime As New MapiProperty(MapiPropertyTag.PR_LAST_MODIFICATION_TIME, ConvertDateTime(New DateTime(2013, 9, 11)))
            mapiMsg.SetProperty(modificationTime)
            mapiMsg.Save(dataDir & Convert.ToString("MapiProp_out.msg"))
            'ExEnd:SetMAPIProperties
        End Sub

        'ExStart:ConvertDateTime
        Private Shared Function ConvertDateTime(t As DateTime) As Byte()
            Dim filetime As Long = t.ToFileTime()
            Dim d As Byte() = New Byte(7) {}
            d(0) = CByte(filetime And &HFF)
            d(1) = CByte((filetime And &HFF00) >> 8)
            d(2) = CByte((filetime And &HFF0000) >> 16)
            d(3) = CByte((filetime And &HFF000000UI) >> 24)
            d(4) = CByte((filetime And &HFF00000000L) >> 32)
            d(5) = CByte((filetime And &HFF0000000000L) >> 40)
            d(6) = CByte((filetime And &HFF000000000000L) >> 48)
            d(7) = CByte((CULng(filetime) And &HFF00000000000000UL) >> 56)
            Return d
        End Function
        'ExEnd:ConvertDateTime
    End Class
End Namespace
