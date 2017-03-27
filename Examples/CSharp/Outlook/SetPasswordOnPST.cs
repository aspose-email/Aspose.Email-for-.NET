using System.IO;
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
    class SetPasswordOnPST
    {
        public static void Run()
        {
            // The path to the File directory.
            // ExStart:SetPasswordOnPST
            string dataDir = RunExamples.GetDataDir_Outlook();

            string alreadyCreated = dataDir + "SetPasswordOnPST_out.pst";
            if (File.Exists(alreadyCreated))
            {
                File.Delete(alreadyCreated);
            }
            else
            {

            }


            using (PersonalStorage pst = PersonalStorage.Create(dataDir + "SetPasswordOnPST_out.pst", FileFormatVersion.Unicode))
            {
                // Set the password
                const string password = "Password1";
                pst.Store.ChangePassword(password);
                // Remove the password
                pst.Store.ChangePassword(null);
            }
            // ExEnd:SetPasswordOnPST
        }
    }
}