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

    Public Class MergePSTFiles
        Shared totalAdded As Integer
        Shared currentFolder As String
        Shared messageCount As Integer

        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()
            Dim dst As String = dataDir & Convert.ToString("Test.pst")

            totalAdded = 0

            Using pst As PersonalStorage = PersonalStorage.FromFile(dst)
                ' The events subscription is an optional step for the tracking process only.
                AddHandler pst.StorageProcessed, AddressOf PstMerge_OnStorageProcessed
                AddHandler pst.ItemMoved, AddressOf PstMerge_OnItemMoved

                ' Merges with the pst files that are located in separate folder. 
                pst.MergeWith(Directory.GetFiles(dataDir & Convert.ToString("chunks\")))
                Console.WriteLine("Total messages added: {0}", totalAdded)
            End Using

            Console.WriteLine(Convert.ToString(Environment.NewLine + "PST merged successfully at ") & dst)
        End Sub

        Private Shared Sub PstMerge_OnStorageProcessed(sender As Object, e As StorageProcessedEventArgs)
            Console.WriteLine("*** The storage is merging: {0}", e.FileName)
        End Sub

        Private Shared Sub PstMerge_OnItemMoved(sender As Object, e As ItemMovedEventArgs)
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
            totalAdded += 1
        End Sub
    End Class
End Namespace
