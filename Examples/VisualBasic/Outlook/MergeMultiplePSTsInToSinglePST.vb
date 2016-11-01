Imports System.IO
Imports Aspose.Email.Outlook.Pst

Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class MergeMultiplePSTsInToSinglePST
        Shared totalAdded As Integer
        Shared currentFolder As String
        Shared messageCount As Integer

        Public Shared Sub Run()
            ' ExStart:MergePSTFiles
            ' The path to the File directory.
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()
            Dim dst As String = dataDir & Convert.ToString("Sub.pst")
            totalAdded = 0
            Try
                Using personalStorage__1 As PersonalStorage = PersonalStorage.FromFile(dst)
                    ' The events subscription is an optional step for the tracking process only.
                    AddHandler personalStorage__1.StorageProcessed, AddressOf PstMerge_OnStorageProcessed
                    AddHandler personalStorage__1.ItemMoved, AddressOf PstMerge_OnItemMoved

                    ' Merges with the pst files that are located in separate folder. 
                    personalStorage__1.MergeWith(Directory.GetFiles(dataDir & Convert.ToString("MergePST\")))
                    Console.WriteLine("Total messages added: {0}", totalAdded)
                End Using
                ' ExEnd:MergePSTFiles
                Console.WriteLine(Convert.ToString(Environment.NewLine + "PST merged successfully at ") & dst)
            Catch ex As Exception
                Console.WriteLine(ex.Message + vbLf & "This example will only work if you apply a valid Aspose Email License. You can purchase full license or get 30 day temporary license from http:// Www.aspose.com/purchase/default.aspx.")
            End Try
        End Sub

        Private Shared Sub PstMerge_OnStorageProcessed(sender As Object, e As StorageProcessedEventArgs)
            Console.WriteLine("*** The storage is merging: {0}", e.FileName)
        End Sub

        Private Shared Sub PstMerge_OnItemMoved(sender As Object, e As ItemMovedEventArgs)
            ' ExStart:MergePSTFiles-PstMerge_OnItemMoved
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
            ' ExEnd:MergePSTFiles-PstMerge_OnItemMoved
        End Sub
    End Class
End Namespace