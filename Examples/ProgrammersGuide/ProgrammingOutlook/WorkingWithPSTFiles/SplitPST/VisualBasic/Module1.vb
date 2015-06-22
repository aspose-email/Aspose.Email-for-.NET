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
    Dim messageCount As Integer

    Sub Main()
        'Initialize license

        'Split the PST
        Dim dataDir As String = Path.GetFullPath("../../../Data/")
        Using pst As PersonalStorage = PersonalStorage.FromFile(dataDir & Convert.ToString("Outlook.pst"))
            ' The events subscription is an optional step for the tracking process only.
            'pst.StorageProcessed += PstSplit_OnStorageProcessed()
            'pst.ItemMoved += PstSplit_OnItemMoved()

            'pst.StorageProcessed = (pst.StorageProcessed + PstSplit_OnStorageProcessed())
            'pst.ItemMoved = (pst.ItemMoved + PstSplit_OnItemMoved())

            ' Splits into pst chunks with the size of 300 KB
            pst.SplitInto(300 * 1024, dataDir & Convert.ToString("Chunks\"))
        End Using

        Console.ReadKey()
    End Sub

    Private Function PstSplit_OnItemMoved(sender As Object, e As ItemMovedEventArgs)
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
    End Function

    Private sub PstSplit_OnStorageProcessed(sender As Object, e As StorageProcessedEventArgs)
        If currentFolder IsNot Nothing Then
            Console.WriteLine("    Added {0} messages to ""{1}""", messageCount, currentFolder)
        End If

        messageCount = 0
        currentFolder = Nothing
        Console.WriteLine("*** The chunk is processed: {0}", e.FileName)
    End Sub

End Module
