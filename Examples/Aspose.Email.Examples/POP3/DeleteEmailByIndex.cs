using Aspose.Email.Clients.Pop3;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Aspose.Email.Examples.POP3
{
    class DeleteEmailByIndex
    {
        public static void Run()
        {
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
        }
    }
}
