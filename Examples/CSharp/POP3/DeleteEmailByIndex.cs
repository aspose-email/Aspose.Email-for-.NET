using Aspose.Email.Clients.Pop3;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.POP3
{
    class DeleteEmailByIndex
    {
        public static void Run()
        {
            // ExStart:DeleteEmailByIndex
            // Create a POP3 client
            Pop3Client client = new Pop3Client("mail.aspose.com", 110, "username", "psw");
            try
            {
                // Delete all the message one by one
                int messageCount = client.GetMessageCount();
                for (int i = 1; i <= messageCount; i++)
                {
                    client.DeleteMessage(i);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            // ExEnd:DeleteEmailByIndex
        }
    }
}
