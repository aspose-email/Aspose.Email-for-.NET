Imports System.IO
Imports Aspose.Email.Mail
Imports Aspose.Email.Outlook
Imports Aspose.Email.Pop3
Imports Aspose.Email
Imports Aspose.Email.Mime
Imports Aspose.Email.Imap
Imports System.Configuration
Imports System.Data
Imports Aspose.Email.Mail.Bounce
Imports Aspose.Email.Exchange
Imports Aspose.Email.Outlook.Pst

Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook

    Public Class NewPSTAddSubfolders
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()
            Dim dst As String = dataDir & Convert.ToString("PersonalStorage.pst")

            If File.Exists(dst) Then
                File.Delete(dst)
            End If

            ' Create new PST
            Dim pst As PersonalStorage = PersonalStorage.Create(dst, FileFormatVersion.Unicode)

            ' Add new folder "Inbox"
            pst.RootFolder.AddSubFolder("Inbox")

            Console.WriteLine(Environment.NewLine + "PST saved successfully at " & dst)
        End Sub
    End Class
End Namespace