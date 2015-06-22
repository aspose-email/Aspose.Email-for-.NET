'''///////////////////////////////////////////////////////////////////////
' Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Email. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'''///////////////////////////////////////////////////////////////////////

Imports Aspose.Email.Outlook.Pst
Imports System.IO

Module Module1

    Dim currentFolder As String
    Dim totalAdded As Integer
    Dim messageCount As Integer

    Sub Main()

    End Sub

    Public Sub Merge()
        totalAdded = 0

        Dim dataDir As String = Path.GetFullPath("../../../Data/")
        Using pst As PersonalStorage = PersonalStorage.FromFile(dataDir & Convert.ToString("Test.pst"))
            ' The events subscription is an optional step for the tracking process only.
            'AddHandler pst.StorageProcessed, AddressOf PstMerge_OnStorageProcessed
            'AddHandler pst.ItemMoved, AddressOf PstMerge_OnItemMoved

            ' Merges with the pst files that are located in separate folder. 
            pst.MergeWith(Directory.GetFiles(dataDir & Convert.ToString("chunks\")))
            Console.WriteLine("Total messages added: {0}", totalAdded)
        End Using
    End Sub

    Private Sub PstMerge_OnStorageProcessed(sender As Object, e As StorageProcessedEventArgs)
        Console.WriteLine("*** The storage is merging: {0}", e.FileName)
    End Sub

    Private Sub PstMerge_OnItemMoved(sender As Object, e As ItemMovedEventArgs)
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
End Module
