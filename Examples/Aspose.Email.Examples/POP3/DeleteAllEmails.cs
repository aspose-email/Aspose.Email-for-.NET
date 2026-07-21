using Aspose.Email.Clients.Pop3;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Aspose.Email.Examples.POP3
{
    class DeleteAllEmails
    {
        public static void Run()
        {           
            // Create a POP3 client
            Pop3Client client = new Pop3Client("mail.aspose.com", 110, "username", "psw");
            try
            {
                // Delete all the messages
                client.DeleteMessages();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
