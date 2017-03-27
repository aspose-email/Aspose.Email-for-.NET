using Aspose.Email.Clients;
using Aspose.Email.Clients.Imap;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

/* This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET 
   API reference when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq 
   for more information. If you do not wish to use NuGet, you can manually download 
   Aspose.Email for .NET API from http://www.aspose.com/downloads, 
   install it and then add its reference to this project. For any issues, questions or suggestions 
   please feel free to contact us using http://www.aspose.com/community/forums/default.aspx            
*/

namespace Aspose.Email.Examples.CSharp.Email.IMAP
{
    class DeletingAndRenamingFolders
    {
        public static void Run()
        {
            // Create an instance of the ImapClient class
            ImapClient client = new ImapClient();

            // Specify host, username and password, and set port for your client
            client.Host = "imap.gmail.com";
            client.Username = "your.username@gmail.com";
            client.Password = "your.password";
            client.Port = 993;
            client.SecurityOptions = SecurityOptions.Auto;
            try
            {
                // ExStart:DeletingAndRenamingFolders
                // Delete a folder and Rename a folder
                client.DeleteFolder("foldername");
                client.RenameFolder("foldername", "newfoldername");
                // ExEnd:DeletingAndRenamingFolders

                client.Dispose();
            }
            catch (Exception ex)
            {
                Console.Write(Environment.NewLine + ex);
            }
        }
    }
}
