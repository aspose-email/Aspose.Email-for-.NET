//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Email. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Email;
using Aspose.Email.Outlook.Pst;

namespace NewPSTAddSubfolders
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            // Create new PST
            PersonalStorage pst = PersonalStorage.Create(dataDir + "PersonalStorage.pst", FileFormatVersion.Unicode);

            // Add new folder "Inbox"
            pst.RootFolder.AddSubFolder("Inbox");

            // Display Status.
            System.Console.WriteLine("New PST file created and addition of folder done successfully in it.");
        }
    }
}