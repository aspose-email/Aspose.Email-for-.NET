using Aspose.Email.Mapi;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class SaveMsgAsTemplate
    {
        static void Run()
        {            
            ///<summary>
            /// This example shows how to save an Outlook MSG file to Outlook Template  using the MapiMessage API
            /// Available from Aspose.Email for .NET 6.4.0 onwards
            /// MapiMessage.SaveAsTemplate(Stream stream) - Saves to the specified stream as Outlook File Template(OFT format).
            /// MapiMessage.SaveAsTemplate(string fileName) - Saves to the specified file as Outlook File Template(OFT format).
            ///</summary>

            // ExStart: SaveMsgAsTemplate
            using (MapiMessage mapi = new MapiMessage("test@from.to", "test@to.to", "template subject", "Template body"))
            {
                string oftMapiFileName = "mapiToOft.msg";
                mapi.SaveAsTemplate(oftMapiFileName);
            }            
            // ExEnd: SaveMsgAsTemplate
        }
    }
}
