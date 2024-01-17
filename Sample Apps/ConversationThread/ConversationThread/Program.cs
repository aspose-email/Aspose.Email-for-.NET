/// This is a simple console example shows how to use Aspose.Email to search conversation related messages in PST.
/// The <see cref="ConversationUtil.GroupMessagesByThread()"/> method traverses the specified folder and returns the <see cref="MessageThread"/> collection.
/// <see cref="MessageThread"/> is a class that contains identifiers of conversation related messages.
/// The <see cref="ConversationUtil.SaveThreads(List{MessageThread}, string)"/> method extracts and saves these messages.
/// Related messages will be in the same folder.

// See https://aka.ms/new-console-template for more information
using Aspose.Email;
using Aspose.Email.Storage.Pst;
using ConversationThread;

// Path to the data.
var dataDirName = @"..\..\..\..\Data";
// The source PST file.
// Put it in the Data folder or set your path.
var inputFileName = "YOUR_PST_FILENAME";
// Path to the output folder.
var outDirectory = Path.Combine(dataDirName, "Out");

var invalidChars = Path.GetInvalidFileNameChars();
var consoleProgress = ProcessAnimation.Create();
var conversations = new List<MessageThread>();

// Set the Aspose.Email license.
try
{
    var lic = new License();
    lic.SetLicense(@"YOUR_LICENSE_FILENAME");
}
catch (InvalidOperationException)
{
    Console.ForegroundColor = ConsoleColor.Red;
    Console.WriteLine("Error: License file not found. Please set your license.");
    Console.ResetColor();
    Environment.Exit(-1);
}

using (var util = new ConversationUtil(
    Path.Combine(dataDirName, inputFileName), 
    StandardIpmFolder.Inbox, 
    delegate { consoleProgress.Write(); }))
{
    try
    {
        Console.Write("Find all threads...");
        consoleProgress = ProcessAnimation.Create();

        // Traverses the specified folder and returns the conversations.
        conversations = util.GroupMessagesByThread();
    }
    catch (Exception ex)
    {
        consoleProgress.Fail();
        Console.WriteLine(ex.ToString());
        Environment.Exit(-1);
    }

    consoleProgress.Success();
    
    Console.WriteLine();
    
    try
    {
        Console.Write("Extract messages...");
        consoleProgress = ProcessAnimation.Create();

        // Extracts and saves these conversation messages.
        util.SaveThreads(conversations, outDirectory);
    }
    catch ( Exception ex )
    {
        consoleProgress.Fail();
        Console.WriteLine(ex.ToString());
        Environment.Exit(-1);
    }
}

consoleProgress.Success();
Console.WriteLine();
Console.WriteLine($"Check the results in {Path.GetFullPath(outDirectory)}");
Console.WriteLine();
