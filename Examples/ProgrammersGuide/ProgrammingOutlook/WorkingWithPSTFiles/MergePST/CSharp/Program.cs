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
namespace MergePSTFiles
{
    class Program
    {
        static int totalAdded;
        static string currentFolder;
        static int messageCount;

        static void Main(string[] args)
        {
            Merge();
        }

        public static void Merge()
        {
            totalAdded = 0;

            string dataDir = Path.GetFullPath("../../../Data/");
            using (PersonalStorage pst = PersonalStorage.FromFile(dataDir + @"Test.pst"))
            {
                // The events subscription is an optional step for the tracking process only.
                pst.StorageProcessed += PstMerge_OnStorageProcessed;
                pst.ItemMoved += PstMerge_OnItemMoved;

                // Merges with the pst files that are located in separate folder. 
                pst.MergeWith(Directory.GetFiles(dataDir + @"chunks\"));
                Console.WriteLine("Total messages added: {0}", totalAdded);
            }
        }

        static void PstMerge_OnStorageProcessed(object sender, StorageProcessedEventArgs e)
        {
            Console.WriteLine("*** The storage is merging: {0}", e.FileName);
        }

        static void PstMerge_OnItemMoved(object sender, ItemMovedEventArgs e)
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
            totalAdded++;
        }


    }
}
