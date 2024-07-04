using Aspose.Email.Clients.Graph;

namespace GraphApp;

public static class FolderHierarchy
{
    
    /// <summary>
    /// Retrieves all folders in the mailbox recursively and returns a hierarchical collection of FolderNode objects.
    /// </summary>
    /// <param name="client">The IGraphClient instance.</param>
    /// <returns>A hierarchical collection of FolderNode objects.</returns>
    public static List<FolderNode> Retrieve(IGraphClient client)
    {
        try
        {
            // Retrieve root folders
            var rootFolders = client.ListFolders();
            var allFolders = new List<FolderNode>();

            // Retrieve subfolders recursively
            foreach (var folder in rootFolders)
            {
                var folderNode = new FolderNode(folder);
                RetrieveSubFolders(client, folderNode);
                allFolders.Add(folderNode);
            }

            return allFolders;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to retrieve folders: {ex.Message}");
            throw;
        }
    }

    /// <summary>
    /// Retrieves subfolders recursively and adds them to the parent FolderNode's SubFolders property.
    /// </summary>
    /// <param name="client">The IGraphClient instance.</param>
    /// <param name="parentFolderNode">The parent FolderNode object.</param>
    private static void RetrieveSubFolders(IGraphClient client, FolderNode parentFolderNode)
    {
        try
        {
            if (parentFolderNode.Folder.HasSubFolders)
            {
                var subFolders = client.ListFolders(parentFolderNode.Folder.ItemId);
                
                foreach (var subFolder in subFolders)
                {
                    var subFolderNode = new FolderNode(subFolder);
                    RetrieveSubFolders(client, subFolderNode);
                    parentFolderNode.SubFolders.Add(subFolderNode);
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to retrieve subfolders for {parentFolderNode.Folder.DisplayName}: {ex.Message}");
            throw;
        }
    }
}


