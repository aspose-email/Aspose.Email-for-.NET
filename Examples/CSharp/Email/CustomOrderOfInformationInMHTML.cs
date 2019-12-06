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
    class CustomOrderOfInformationInMHTML
    {
        public static void Run()
        {
            // ExStart:1
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Email();

            MailMessage eml = MailMessage.Load(dataDir + "Attachments.eml");
            MhtSaveOptions opt = SaveOptions.DefaultMhtml;

            eml.Save(dataDir + "CustomOrderOfInformationInMHTML_1.mhtml", opt);

            opt.RenderingHeaders.Add(MhtTemplateName.From);
            opt.RenderingHeaders.Add(MhtTemplateName.Subject);
            opt.RenderingHeaders.Add(MhtTemplateName.To);
            opt.RenderingHeaders.Add(MhtTemplateName.Sent);

            eml.Save(dataDir + "CustomOrderOfInformationInMHTML_2.mhtml", opt);

            opt.RenderingHeaders.Clear();
            opt.RenderingHeaders.Add(MhtTemplateName.Attachments);
            opt.RenderingHeaders.Add(MhtTemplateName.Cc);
            opt.RenderingHeaders.Add(MhtTemplateName.Subject);

            eml.Save(dataDir + "CustomOrderOfInformationInMHTML_3.mhtml", opt);
            // ExEnd:1
            Console.WriteLine("CustomOrderOfInformationInMHTML executed successfully");
        }
    }
}
