using System;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    class IMAP4ExtendedListCommand
    {
        public static void Run()
        {
            using (ImapClient client = new ImapClient("imap.gmail.com", 993, "username", "password"))
            {
                ImapFolderInfoCollection folderInfoCol = client.ListFolders("*");
                Console.WriteLine("Extended List Supported: " + client.ExtendedListSupported);
                foreach (ImapFolderInfo folderInfo in folderInfoCol)
                {
                    switch (folderInfo.Name)
                    {
                        case "[Gmail]/All Mail":
                            Console.WriteLine("Has Children: " + folderInfo.HasChildren);
                            break;
                        case "[Gmail]/Bin":
                            Console.WriteLine("Bin has children? " + folderInfo.HasChildren);
                            break;
                        case "[Gmail]/Drafts":
                            Console.WriteLine("Drafts has children? " + folderInfo.HasChildren);
                            break;
                        case "[Gmail]/Important":
                            Console.WriteLine("Important has Children? " + folderInfo.HasChildren);
                            break;
                        case "[Gmail]/Sent Mail":
                            Console.WriteLine("Sent Mail has Children? " + folderInfo.HasChildren);
                            break;
                        case "[Gmail]/Spam":
                            Console.WriteLine("Spam has Children? " + folderInfo.HasChildren);
                            break;
                        case "[Gmail]/Starred":
                            Console.WriteLine("Starred has Children? " + folderInfo.HasChildren);
                            break;
                    }
                }
            }
        }
    }
}
