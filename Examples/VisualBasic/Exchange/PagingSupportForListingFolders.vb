Imports Aspose.Email.Exchange

Namespace Aspose.Email.Examples.VisualBasic.Email.Exchange

    Public Class PagingSupportForListingFolders
        Private Shared Sub Run()
            'ExStart: PagingSupportForListingFolders
            '<summary>
            'This example shows how to retrieve folders information from Exchange Server with Paging Support
            'Introduced in Aspose.Email for .NET 6.4.0
            '</summary>
            Using client As IEWSClient = EWSClient.GetEWSClient("exchange.domain.com", "username", "password")
                Dim itemsPerPage As Integer = 5
                Dim totalFoldersCollection As ExchangeFolderInfoCollection = client.ListSubFolders(client.MailboxInfo.RootUri)
                Console.WriteLine(totalFoldersCollection.Count)

                '/ RETREIVING INFORMATION USING PAGING SUPPORT 

                Dim pages As New List(Of ExchangeFolderPageInfo)()
                Dim pagedFoldersCollection As ExchangeFolderPageInfo = client.ListSubFoldersByPage(client.MailboxInfo.RootUri, itemsPerPage)

                Console.WriteLine(pagedFoldersCollection.TotalCount)

                pages.Add(pagedFoldersCollection)
                While Not pagedFoldersCollection.LastPage
                    pagedFoldersCollection = client.ListSubFoldersByPage(client.MailboxInfo.RootUri, itemsPerPage, pagedFoldersCollection.PageOffset + 1)
                    pages.Add(pagedFoldersCollection)
                End While
                Dim retrievedFolders As Integer = 0
                For Each pageCol As ExchangeFolderPageInfo In pages
                    retrievedFolders += pageCol.Items.Count
                Next

                'Verify the total count of folders
                Console.WriteLine(retrievedFolders)
            End Using
            'ExEnd: PagingSupportForListingFolders
        End Sub
    End Class
End Namespace