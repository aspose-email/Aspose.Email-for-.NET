//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Email. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.IO;
using System;
using Aspose.Email;
using Aspose.Email.Outlook.Pst;

namespace CSharp.Outlook
{
    class NewPSTAddSubfolders
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Outlook();
            string dst = dataDir + "PersonalStorage.pst";

            if (File.Exists(dst))
                File.Delete(dst);
            
            // Create new PST
            PersonalStorage pst = PersonalStorage.Create(dst, FileFormatVersion.Unicode);

            // Add new folder "Inbox"
            pst.RootFolder.AddSubFolder("Inbox");

            Console.WriteLine(Environment.NewLine + "PST saved successfully at " + dst);
        }
    }
}
