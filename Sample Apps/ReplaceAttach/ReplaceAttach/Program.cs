// This code sample shows how to replace MapiMessage attachments by index.

using Aspose.Email;
using Aspose.Email.Mapi;
using System.Text.RegularExpressions;

// See https://aka.ms/new-console-template for more information

var dataDirName = @"..\..\..\..\Data";
var inputFileName = "Fwd_ This is test mail for outlook attachment.msg";
var attachDirName = Path.Combine(dataDirName, "Attachments");

// Set the license.
// Evaluation version is limited to processing no more than 3 attachments.
try
{ 
    var lic = new License();
    lic.SetLicense(@"YOU_LICENSE_FILENAME"); 
}
catch (InvalidOperationException)
{ 
    Console.ForegroundColor = ConsoleColor.Red;
    Console.WriteLine("Warning: License file not found. Please set your license.");
    Console.ForegroundColor = ConsoleColor.White;
    Console.WriteLine();
    Console.WriteLine("Press <ENTER> if you want to continue with the evaluation version.");
    Console.WriteLine("NOTE: Evaluation version is limited to processing no more than 3 attachments.");
    
    while (Console.ReadKey().Key != ConsoleKey.Enter) { return; }
}

// Load the MSG
var msg = MapiMessage.Load(Path.Combine(dataDirName, inputFileName));

// This list is for mapping old attach names to new names.
var newAttachNames = new List<string>();

// Find the index of an existing attachment and map it to the index of the new attachment.
// Filling the list.
foreach (var attachment in msg.Attachments)
{
    // Retrieve the index of the current attachment.
    // It is a very simplified regex template. But you can replace it with your own.
    var index = Regex.Match(attachment.FileName, @"\d+").Value;

    // Get the name of the new attachment.
    var newAttachName = $"attach{index}.zip";
    newAttachNames.Add(newAttachName);
}

// Removing old attachments.
msg.Attachments.Clear();

// Adding new attachments.
foreach (var attachName in newAttachNames)
{
    // Read the attachment data from the file.
    var bytes = File.ReadAllBytes(Path.Combine(attachDirName, attachName));
    
    msg.Attachments.Add(attachName, bytes);
}

// Save the modified message to a new file.
msg.Save(@Path.Combine(dataDirName, "output.msg"));

Console.ForegroundColor = ConsoleColor.Green;
Console.WriteLine("Done! Check an output file in the Data folder.");
Console.ForegroundColor = ConsoleColor.White;
