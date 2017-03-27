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
    class CreateAndSaveDistributionList
    {
        public static void Run()
        {
            // ExStart:CreateAndSaveDistributionList
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Outlook();

            // Create MapiDistributionListMemberCollection object and add MapiDistributionListMembers
            MapiDistributionListMemberCollection oneOffmembers = new MapiDistributionListMemberCollection();
            oneOffmembers.Add(new MapiDistributionListMember("John R. Patrick", "JohnRPatrick@armyspy.com"));
            oneOffmembers.Add(new MapiDistributionListMember("Tilly Bates", "TillyBates@armyspy.com"));

            MapiDistributionList dlist = new MapiDistributionList("Simple list", oneOffmembers);
            // Set MapiDistributionList properties
            dlist.Body = "Test body";
            dlist.Subject = "Test subject";
            dlist.Mileage = "Test mileage";
            dlist.Billing = "Test billing";

            dlist.Save(dataDir + @"distlist_out.msg");
            // ExEnd:CreateAndSaveDistributionList
        }
    }
}