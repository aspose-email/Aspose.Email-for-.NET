using System;
using Aspose.Email.Mapi;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class ReadNamedMAPIProperties
    {
        public static void Run()
        {
            //ExStart:ReadNamedMAPIProperties
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Outlook();

            // Load from file
            MapiMessage message = MapiMessage.FromFile(dataDir + @"message.msg");

            // Get all named MAPI properties
            MapiPropertyCollection properties = message.NamedProperties;
            // Read all properties in foreach loop
            foreach (MapiNamedProperty mapiNamedProp in properties.Values)
            {
                // Read any specific property
                switch (mapiNamedProp.NameId)
                {
                    case "TEST":
                        Console.WriteLine("{0} = {1}", mapiNamedProp.NameId, mapiNamedProp.GetString());
                        break;
                    case "MYPROP":
                        Console.WriteLine("{0} = {1}", mapiNamedProp.NameId, mapiNamedProp.GetString());
                        break;
                    default: break;
                }
            }
            //ExEnd:ReadNamedMAPIProperties
        }
    }
}