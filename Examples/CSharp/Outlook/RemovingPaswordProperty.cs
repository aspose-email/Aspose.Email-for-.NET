using System;
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
    class RemovingPaswordProperty
    {
        public static void Run()
        {
            // The path to the File directory.
            // ExStart:RemovingPaswordProperty
            string dataDir = RunExamples.GetDataDir_Outlook();
            PersonalStorage personalStorage = PersonalStorage.FromFile(dataDir + "PersonalStorage1.pst");
            if (personalStorage.Store.Properties.ContainsKey(MapiPropertyTag.PR_PST_PASSWORD))
            {
                MapiProperty property = new MapiProperty(MapiPropertyTag.PR_PST_PASSWORD, BitConverter.GetBytes((long)0));
                personalStorage.Store.SetProperty(property);
            }
            // ExEnd:RemovingPaswordProperty
        }
    }
}