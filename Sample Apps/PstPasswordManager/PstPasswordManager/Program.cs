/*-------------------------------------------------------------------------------------------
A console app example for managing password-protected PST files. 
This example demonstrates some features of Aspose.Email in action.

Functionality:
  - Accepts command-line arguments to specify the PST file path and the desired action.
  - Validates the provided password against the password-protected PST file.
  - Adds a password to the PST file.
  - Changes the existing password on the PST file to a new one.
  - Removes the password from the PST file.
  
To use the app, follow these steps:
- Navigate to the App Directory: Open a terminal or command prompt and navigate to the directory where the application is located.
- Run the Application with Command-Line Arguments. For Example: 'dotnet run -- <PST file path> -v "[password string]"' 
---------------------------------------------------------------------------------------------*/

using Aspose.Email.Storage.Pst;

// Check if no arguments are provided
if (args.Length == 0)
{
    // Show usage information
    ShowUsage();
    return;
}

string pstFilePath = null; // PST file path
string password = null;    // password if provided
string action = null;      // action to be performed

// Parse command-line arguments
for (var i = 0; i < args.Length; i++)
{
    switch (args[i])
    {
        case "-v":
        case "-a":
        case "-c":
            // Check if the next argument exists and is the password
            if (i + 1 < args.Length)
            {
                action = args[i]; // Set action
                password = args[i + 1]; // Set password
                i++; // Skip the password argument in the loop
            }
            else
            {
                Console.WriteLine($"Error: {args[i]} requires a password.");
                return;
            }
            break;
        case "-r":
            // Set action to remove password
            action = args[i];
            break;
        default:
            if (pstFilePath != null)
            {
                Console.WriteLine($"Error: Unexpected argument {args[i]}.");
                return;
            }
            
            // Assume the first unrecognized argument is the PST file path
            // Set PST file path
            pstFilePath = args[i];
            break;
    }
}

// Check if PST file path is provided
if (pstFilePath == null)
{
    Console.WriteLine("Error: PST file path is required.");
    ShowUsage();
    return;
}

ExecuteAction(pstFilePath, action, password);

// Display usage information
static void ShowUsage()
{
    Console.WriteLine("Usage:");
    Console.WriteLine("  PstPasswordManager <PST file path> [-v \"[password string]\"] [-a \"[password string]\"] [-c \"[password string]\"] [-r]");
    Console.WriteLine("Options:");
    Console.WriteLine("  -v \"[password string]\"    Validate password in password protected PST");
    Console.WriteLine("  -a \"[password string]\"    Add password on PST file");
    Console.WriteLine("  -c \"[password string]\"    Change password on PST file");
    Console.WriteLine("  -r                        Remove password on PST file");
}

// Execute the desired action on the PST file
static void ExecuteAction(string pstFilePath, string action, string password)
{
    try
    {
        // Open a PST file
        using var pst = PersonalStorage.FromFile(pstFilePath);
        
        switch (action)
        {
            case "-v":
                // Validate password
                if (!pst.Store.IsPasswordProtected)
                {
                    Console.WriteLine("The PST file is not password protected. First add a password.");
                    break;
                }
                
                Console.WriteLine(pst.Store.IsPasswordValid(password) ? "Password is valid." : "Password is invalid.");
                break;
            case "-a":
                // Add password
                pst.Store.ChangePassword(password);
                Console.WriteLine("Password added successfully.");
                break;
            case "-c":
                // Change password
                if (!pst.Store.IsPasswordProtected)
                {
                    Console.WriteLine("The PST file is not password protected. First add a password.");
                    break;
                }

                pst.Store.ChangePassword(password);
                Console.WriteLine("Password changed successfully.");
                break;
            case "-r":
                // Remove password
                if (!pst.Store.IsPasswordProtected)
                {
                    Console.WriteLine("The PST file is not password protected anyway.");
                    break;
                }

                pst.Store.ChangePassword(null);
                Console.WriteLine("Password removed successfully.");
                break;
            default:
                // Check if the PST file is password protected
                Console.WriteLine(pst.Store.IsPasswordProtected
                    ? "The PST file is password protected."
                    : "The PST file is not password protected.");
                break;
        }
    }
    catch (Exception ex)
    {
        // Handle any exceptions that occur during PST file operations
        Console.WriteLine($"Error: {ex.Message}");
    }
}