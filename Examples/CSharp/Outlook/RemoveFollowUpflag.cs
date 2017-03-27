using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Email;
using Aspose.Email.Examples.CSharp.Email;
using Aspose.Email.Mime;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class RemoveFollowUpflag
    {
        public static void Run()
        {
            //ExStart:RemoveFollowUpflag
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Outlook();

            // Load file from Disk
            MapiMessage mapi = MapiMessage.FromFile(dataDir + "message.msg");
            FollowUpManager.ClearFlag(mapi);
            mapi.Save(dataDir + "RemoveFollowUpflag_out.msg");
            //ExEnd:RemoveFollowUpflag
        }
    }
}