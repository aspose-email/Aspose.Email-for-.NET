using Aspose.Email.Storage.Pst;

/* This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET 
   API reference when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq 
   for more information. If you do not wish to use NuGet, you can manually download 
   Aspose.Email for .NET API from http://www.aspose.com/downloads, 
   install it and then add its reference to this project. For any issues, questions or suggestions 
   please feel free to contact us using http://www.aspose.com/community/forums/default.aspx            
*/


namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class ConvertOSTToPST
    {
        public static void Run()
        {
            // ExStart:ConvertOSTToPST
            string dataDir = RunExamples.GetDataDir_Outlook();

            // Load the Outlook file
            string path = dataDir + "PersonalStorage.pst";

            // Get the Display Name of the file
            using (PersonalStorage ost = PersonalStorage.FromFile(path))
            {
                ost.SaveAs(dataDir + "ConvertOSTToPST_out.pst", FileFormat.Pst);
            }
            // ExEnd:ConvertOSTToPST
        }
    }
}