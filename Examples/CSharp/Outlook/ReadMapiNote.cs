﻿using Aspose.Email.Mapi;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from https://www.nuget.org/packages/Aspose.Email/, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using https://forum.aspose.com/c/email
*/

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class ReadMapiNote
    {
        public static void Run()
        {
            // ExStart:ReadMapiNote
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Outlook();
            MapiMessage note = MapiMessage.FromFile(dataDir + "MapiNote.msg");
            MapiNote note2 = (MapiNote)note.ToMapiMessageItem();
            // ExEnd:ReadMapiNote
        }
    }
}