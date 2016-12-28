Imports System.IO
Imports Aspose.Email.Outlook.Pst

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'


Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class SetPasswordOnPST
        Public Shared Sub Run()
            ' The path to the documents directory.
            ' ExStart:SetPasswordOnPST
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()

            Dim checkfile As String = dataDir + "SetPasswordOnPST_out.pst"

            If File.Exists(checkfile) Then
                File.Delete(checkfile)

            Else
            End If

            Using pst As PersonalStorage = PersonalStorage.Create(dataDir + "SetPasswordOnPST_out.pst", FileFormatVersion.Unicode)
                ' Set the password
                Const password As String = "Password1"
                pst.Store.ChangePassword(password)
                ' Remove the password
                pst.Store.ChangePassword(Nothing)
            End Using
            ' ExEnd:SetPasswordOnPST
        End Sub
    End Class
End Namespace
