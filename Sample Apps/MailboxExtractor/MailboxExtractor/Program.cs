/*
 * This sample program uses the Aspose.Email library to extract messages from mailbox storage files
 * and save them to a specified output directory.
 * It validates the format of the input file, creates the output directory if it doesn't exist, and handles any exceptions 
 * that occur during the extraction process.
 * Supported Storage File Formats: PST/OST, MBOX, OLM, TGZ
 */

using MailboxExtractor;  

// Check if the correct number of arguments is provided
if (args.Length != 2)
{
    // Display usage instructions if the number of arguments is incorrect
    Console.WriteLine("Usage: MailboxExtractor <PST file path> <Output directory>");
    return;
}

var fileName = args[0];    // The first argument is the path to the storage file
var outputDir = args[1];   // The second argument is the output directory

// Check if the provided storage file exists
if (!File.Exists(fileName))
{
    // Display an error message if the PST file does not exist
    Console.WriteLine($"Error: PST file '{fileName}' not found.");
    return;
}

// Create the output directory if it doesn't exist
if (!Directory.Exists(outputDir))
{
    // Create the output directory
    Directory.CreateDirectory(outputDir);
}

try
{
    var mailbox = Mailbox.Open(fileName);  // Open the storage file
    Console.Write("Extract messages...");
    mailbox.SaveTo(outputDir);  // Save the extracted messages to the output directory
}
catch (Exception ex)
{
    // Display an error message and exit if an exception occurs
    Console.WriteLine($"Error: {ex.Message}");
    Environment.Exit(-1);
}
