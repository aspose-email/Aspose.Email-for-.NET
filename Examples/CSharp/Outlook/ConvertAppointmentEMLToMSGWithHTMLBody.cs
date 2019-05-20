using Aspose.Email.Mapi;
using Aspose.Email.PersonalInfo.VCard;
using System;

/* This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET 
   API reference when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq 
   for more information. If you do not wish to use NuGet, you can manually download 
   Aspose.Email for .NET API from https://www.nuget.org/packages/Aspose.Email/, 
   install it and then add its reference to this project. For any issues, questions or suggestions 
   please feel free to contact us using https://forum.aspose.com/c/email            
*/

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class ConvertAppointmentEMLToMSGWithHTMLBody
    {
        public static void Run()
        {
            // ExStart:1
            // Outlook directory
            string dataDir = RunExamples.GetDataDir_Outlook();

            MailMessage mailMessage = MailMessage.Load(dataDir + "TestAppointment.eml");

            MapiConversionOptions conversionOptions = new MapiConversionOptions();
            conversionOptions.Format = OutlookMessageFormat.Unicode;

            // default value for ForcedRtfBodyForAppointment is true
            conversionOptions.ForcedRtfBodyForAppointment = false;

            MapiMessage mapiMessage = MapiMessage.FromMailMessage(mailMessage, conversionOptions);

            Console.WriteLine("Body Type: " + mapiMessage.BodyType);

            mapiMessage.Save(dataDir + "TestAppointment_out.msg");
            // ExEnd:1

            Console.WriteLine("ConvertAppointmentEMLToMSGWithHTMLBody executed successfully.");
        }
    }
}