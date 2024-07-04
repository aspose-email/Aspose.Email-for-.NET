using Aspose.Email.Clients.Graph;

namespace GraphApp;

/// <summary>
/// Represents a node in a folder hierarchy,
/// extending the properties of FolderInfo and storing a collection of subfolders.
/// </summary>
public class FolderNode
{
    /// <summary>
    /// Gets the FolderInfo object representing the folder information.
    /// </summary>
    public FolderInfo Folder { get; }
    
    /// <summary>
    /// Gets he collection of subfolders contained within the current folder.
    /// </summary>
    public List<FolderNode?> SubFolders { get; }

    /// <summary>
    /// Initializes a new instance of the FolderNode class with the specified FolderInfo object.
    /// </summary>
    /// <param name="folder">The FolderInfo object representing the folder information.</param>
    public FolderNode(FolderInfo folder)
    {
        Folder = folder;
        SubFolders = new List<FolderNode?>();
    }
    
    /// <summary>
    /// Prints all folders in a hierarchical manner starting from the current node.
    /// </summary>
    public void PrintHierarchy()
    {
        PrintFolderNode(this, 0);
    }

    private void PrintFolderNode(FolderNode node, int indentLevel)
    {
        // Print current folder node
        Console.WriteLine($"{new string(' ', indentLevel * 2)}{node}");

        // Recursively print subfolders
        foreach (var subFolder in node.SubFolders)
        {
            PrintFolderNode(subFolder, indentLevel + 1);
        }
    }

    /// <summary>
    /// Get the folder Display Name.
    /// </summary>
    /// <returns></returns>
    public override string ToString()
    {
        return $"{Folder.DisplayName} ({Folder.ContentCount})";
    }
}