Imports System.IO
Imports Aspose.Email.Outlook.Pst

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class MergeFolderFromAnotherPSTFile
        Public Shared totalAdded As Integer = 0
        Public Shared messageCount As Integer
        Public Shared currentFolder As String
        Public Shared Sub Run()
            ' The path to the File directory.
            ' ExStart:MergeMultiplePST
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()
            Try               
                Using pst As PersonalStorage = PersonalStorage.FromFile(dataDir & Convert.ToString("PersonalStorage.pst"))
                    ' The events subscription is an optional step for the tracking process only.
                    AddHandler pst.StorageProcessed, AddressOf PstMerge_OnStorageProcessed
                    AddHandler pst.ItemMoved, AddressOf PstMerge_OnItemMoved

                    ' Merges with the pst files that are located in separate folder.
                    pst.MergeWith(Directory.GetFiles(dataDir & Convert.ToString("\Sources\")))
                    Console.WriteLine("Total messages added: {0}", totalAdded)
                End Using
            Catch ex As Exception
                Console.WriteLine(ex.Message + vbLf & "This example will only work if you apply a valid Aspose Email License. You can purchase full license or get 30 day temporary license from http:// Www.aspose.com/purchase/default.aspx.")
            End Try
            ' ExEnd:MergeMultiplePST
        End Sub

        Private Shared Sub PstMerge_OnStorageProcessed(sender As Object, e As StorageProcessedEventArgs)
            Console.WriteLine("*** The storage is merging: {0}", e.FileName)
        End Sub

        Private Shared Sub PstMerge_OnItemMoved(sender As Object, e As ItemMovedEventArgs)
            Dim folderPath As String = e.DestinationFolder.RetrieveFullPath()

            If currentFolder Is Nothing Then
                currentFolder = e.DestinationFolder.RetrieveFullPath()
            End If

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
