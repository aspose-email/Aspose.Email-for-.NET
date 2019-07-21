using Aspose.Email.Calendar;
using System;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from https://www.nuget.org/packages/Aspose.Email/, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using https://forum.aspose.com/c/email
*/

namespace Aspose.Email.Examples.CSharp.Email
{
    class ChangeProdIdOfICS
    {
        public static void Run()
        {
            // ExStart:1
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Email();

            string description = "Test Description";
            Appointment app = new Appointment("location", "test appointment", description, DateTime.Today,
            DateTime.Today.AddDays(1), "first@test.com", "second@test.com");

            IcsSaveOptions saveOptions = IcsSaveOptions.Default;
            saveOptions.ProductId = "Test Corporation";
            app.Save(dataDir + "ChangeProdIdOfICS.ics", saveOptions);
            // ExEnd:1
        }
    }
}
