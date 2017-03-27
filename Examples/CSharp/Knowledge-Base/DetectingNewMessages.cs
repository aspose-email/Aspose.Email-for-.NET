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

namespace Aspose.Email.Examples.CSharp.Email.Knowledge.Base
{
    class DetectingNewMessages
    {
        // ExStart:DetectingNewMessages
        public static void Run()
        {
            try
            {
                // Connect to the POP3 mail server and check messages.
                Pop3Client pop3Client = new Pop3Client("pop.domain.com", 993, "username", "password");

                // List all the messages
                Pop3MessageInfoCollection msgList = pop3Client.ListMessages();
                foreach (Pop3MessageInfo msgInfo in msgList)
                {
                    // Get the POP3 message's unique ID
                    string strUniqueID = msgInfo.UniqueId;

                    // Search your local database or data store on the unique ID. If a match is found, that means it's already downloaded. Otherwise download and save it.
                    if (SearchPop3MsgInLocalDB(strUniqueID) == true)
                    {
                        // The message is already in the database. Nothing to do with this message. Go to next message.
                    }
                    else
                    {
                        // Save the message
                        SavePop3MsgInLocalDB(msgInfo);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            
        }

        private static void SavePop3MsgInLocalDB(Pop3MessageInfo msgInfo)
        {
            // Open the database connection according to your database. Use public properties (for example msgInfo.Subject) and store in database,
            // for example, " INSERT INTO POP3Mails (UniqueID, Subject) VALUES ('" + msgInfo.UniqueID + "' , '" + msgInfo.Subject + "') and Run the query to store in database.
        }

        private static bool SearchPop3MsgInLocalDB(string strUniqueID)
        {
            // Open the database connection according to your database. Use strUniqueID in the search query to find existing records,
            // for example, " SELECT COUNT(*) FROM POP3Mails WHERE UniqueID = '" + strUniqueID + "'  Run the query, return true if count == 1. Return false if count == 0.
            return false;
        }
        // ExEnd:DetectingNewMessages
    }
}
