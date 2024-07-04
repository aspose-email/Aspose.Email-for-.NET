// This sample uses Aspose.Email.Clients.Graph to interact with Microsoft Graph API.
// It reads configuration from a JSON file, authenticates using OAuth2, retrieves the folder hierarchy
// for a specified user, lists messages in the "Inbox" folder, and saves the first 10 messages to disk.
//
// Ensure to provide your configuration in the appsettings.json file
// and update the inboxName and outputPath constants if needed.

using Aspose.Email;
using Aspose.Email.Clients.Graph;
using GraphApp;

const string inboxName = "Inbox";
const string outputPath = "";

try
{
    // Acquiring an access token and client initialization.
    var config = AuthenticationConfig.ReadFromJsonFile("appsettings.json");
    var tokenProvider = new GraphTokenProvider(config);

    using var client = GraphClient.GetClient(tokenProvider, config.TenantId);
    client.Resource = ResourceType.Users;
    client.ResourceId = config.UserId;
    
    Console.WriteLine();
    Console.WriteLine("Get a folder hierarchy...");
    
    // Get a folder hierarchy
    var folderNodes = FolderHierarchy.Retrieve(client);
    
    // Print folders in a hierarchical order.
    foreach (var folderNode in folderNodes)
    {
        folderNode.PrintHierarchy();
    }
    
    // Get the Inbox folder by name
    var inbox = folderNodes.FirstOrDefault(
        folderNode => folderNode.Folder.DisplayName.Equals(inboxName, StringComparison.OrdinalIgnoreCase))
        ?.Folder;
    
    Console.WriteLine();
    Console.WriteLine("List messages from a folder...");
    
    if (inbox == null)
    {
        throw new Exception($"Folder {inboxName} not found");
    }
    
    // Call client method to list messages in the specified folder
    var messageInfoCollection = client.ListMessages(inbox.ItemId);

    Console.WriteLine();
    Console.WriteLine($"{inboxName}:");

    // Print messages
    foreach (var messageInfo in messageInfoCollection)
    {
        Console.WriteLine($"     - {messageInfo.Subject}");
    }
    
    Console.WriteLine();
    Console.WriteLine("Fetch  and save the first 10 messages...");
    
    List<MessageInfo> firstTenItems = messageInfoCollection.Take(10).ToList();

    var count = 0;
    
    // Fetch  and save the first 10 messages.
    foreach (var messageInfo in firstTenItems)
    {
        var msg = client.FetchMessage(messageInfo.ItemId);
        msg.Save(Path.Combine(outputPath, $"{++count}-{ToValidFileName(msg.Subject)}.msg"), SaveOptions.DefaultMsgUnicode);
    }
    
    Console.WriteLine();
    Console.WriteLine($"Messages have been successfully saved at {outputPath}");
}
catch (Exception ex)
{
    Console.WriteLine(ex.Message);
}

static string ToValidFileName(string name)
{
    string invalidChars = new string(Path.GetInvalidFileNameChars()) + new string(Path.GetInvalidPathChars());
    foreach (char c in invalidChars)
    {
        name = name.Replace(c.ToString(), "_");
    }
    return name;
}
