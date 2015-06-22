//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Email. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Email.Outlook.Pst;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

// This feature is available from Aspose.Email for .NET 5.4.0 onwards
namespace SplitPST
{
    class Program
    {
        private static string currentFolder;
        private static int messageCount;

        static void Main(string[] args)
        {

            //Initialize license

            //Split the PST
            string dataDir = Path.GetFullPath("../../../Data/");
            using (PersonalStorage pst = PersonalStorage.FromFile(dataDir + "Outlook.pst"))
            {
                // The events subscription is an optional step for the tracking process only.
                pst.StorageProcessed += PstSplit_OnStorageProcessed;
                pst.ItemMoved += PstSplit_OnItemMoved;

                // Splits into pst chunks with the size of 300 KB
                pst.SplitInto(300*1024, dataDir + @"Chunks\");
            }

            Console.ReadKey();
        }

        static void PstSplit_OnItemMoved(object sender, ItemMovedEventArgs e)
        {
            if (currentFolder == null)
            {
                currentFolder = e.DestinationFolder.RetrieveFullPath();
            }

            string folderPath = e.DestinationFolder.RetrieveFullPath();

            if (currentFolder != folderPath)
            {
                Console.WriteLine("    Added {0} messages to \"{1}\"", messageCount, currentFolder);
                messageCount = 0;
                currentFolder = folderPath;
            }

            messageCount++;
        }

        static void PstSplit_OnStorageProcessed(object sender, StorageProcessedEventArgs e)
        {
            if (currentFolder != null)
            {
                Console.WriteLine("    Added {0} messages to \"{1}\"", messageCount, currentFolder);
            }

            messageCount = 0;
            currentFolder = null;
            Console.WriteLine("*** The chunk is processed: {0}", e.FileName);
        }



    }
}
