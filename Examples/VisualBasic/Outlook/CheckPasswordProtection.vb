Imports Aspose.Email.Outlook
Imports Aspose.Email.Outlook.Pst

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class CheckPasswordProtection
        Public Shared Sub Run()
            ' ExStart:CheckPasswordProtection
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()

            Dim mapiMessage__1 As MapiMessage = MapiMessage.FromFile(dataDir & Convert.ToString("message1.msg"))
            FollowUpManager.MarkAsCompleted(mapiMessage__1)
            mapiMessage__1.Save(dataDir & Convert.ToString("MarkedCompleted_out.msg"))
        End Sub

        ' <summary>
        ' Determines whether the specified PST is password protected.
        ' </summary>
        Private Shared Function IsPasswordProtected(pst As PersonalStorage) As Boolean
            ' ExStart:CheckPasswordProtection-IsPasswordProtected
            ' If the property exists and is nonzero, then the PST file is password protected.
            If pst.Store.Properties.Contains(MapiPropertyTag.PR_PST_PASSWORD) Then
                Dim passwordHash As Long = pst.Store.Properties(MapiPropertyTag.PR_PST_PASSWORD).GetLong()
                Return passwordHash <> 0
            End If
            Return False
            ' ExEnd:CheckPasswordProtection-IsPasswordProtected
        End Function
    End Class
End Namespace
