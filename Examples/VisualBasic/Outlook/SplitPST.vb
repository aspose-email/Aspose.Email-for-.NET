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

Namespace Aspose.Email.Examples.VisualBasic.Knowledge.Outlook
    Public Class SplitPST
        Private Shared currentFolder As String
        Private Shared messageCount As Integer

        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()
            Dim dst As String = dataDir & Convert.ToString("Outlook.pst")
            Dim dstSplit As [String] = dataDir & Convert.ToString("Chunks-Split\")

            ' Delete the files if already present
            For Each file__1 As String In Directory.GetFiles(dstSplit)
                File.Delete(file__1)
            Next

            Using pst As PersonalStorage = PersonalStorage.FromFile(dst)
                ' The events subscription is an optional step for the tracking process only.
                AddHandler pst.StorageProcessed, AddressOf PstSplit_OnStorageProcessed
                AddHandler pst.ItemMoved, AddressOf PstSplit_OnItemMoved

                ' Splits into pst chunks with the size of 300 KB
                pst.SplitInto(300 * 1024, dstSplit)
            End Using

            Console.WriteLine(Environment.NewLine + "PST split successfully at " & dst)
        End Sub

        Private Shared Sub PstSplit_OnItemMoved(sender As Object, e As ItemMovedEventArgs)
            If currentFolder Is Nothing Then
                currentFolder = e.DestinationFolder.RetrieveFullPath()
            End If

            Dim folderPath As String = e.DestinationFolder.RetrieveFullPath()

            If currentFolder <> folderPath Then
                Console.WriteLine("    Added {0} messages to ""{1}""", messageCount, currentFolder)
                messageCount = 0
                currentFolder = folderPath
            End If

            messageCount += 1
        End Sub

        Private Shared Sub PstSplit_OnStorageProcessed(sender As Object, e As StorageProcessedEventArgs)
            If currentFolder IsNot Nothing Then
                Console.WriteLine("    Added {0} messages to ""{1}""", messageCount, currentFolder)
            End If

            messageCount = 0
            currentFolder = Nothing
            Console.WriteLine("*** The chunk is processed: {0}", e.FileName)
        End Sub
    End Class
End Namespace