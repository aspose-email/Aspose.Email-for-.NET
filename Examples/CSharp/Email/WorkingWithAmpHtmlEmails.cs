using Aspose.Email.Amp;
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
    class WorkingWithAmpHtmlEmails
    {
        public static void Run()
        {
            // ExStart:1
            string dataDir = RunExamples.GetDataDir_Output();

            AmpMessage msg = new AmpMessage();
            msg.HtmlBody = "<html><body> Hello AMP </body></html>";

            //add AmpAnim component
            AmpAnim anim = new AmpAnim(800, 400);
            anim.Src = "https://placekitten.com/800/400";
            anim.Alt = "Test alt";
            anim.Attribution = "The Go gopher was designed by Reneee French";
            anim.Attributes.Layout = LayoutType.Responsive;
            anim.Fallback = "offline";
            msg.AddAmpComponent(anim);

            //add AmpImage component
            AmpImage img = new AmpImage(800, 400);
            img.Src = "https://placekitten.com/800/400";
            img.Alt = "Test alt";
            img.Attributes.Layout = LayoutType.Responsive;
            msg.AddAmpComponent(img);

            //add AmpCarousel component
            AmpCarousel car = new AmpCarousel(800, 400);
            img = new AmpImage(800, 400);
            img.Src = "https://amp.dev/static/img/docs/tutorials/firstemail/photo_by_caleb_woods.jpg";
            img.Alt = "Test 2 alt";
            img.Attributes.Layout = LayoutType.Fixed;
            car.Images.Add(img);
            img = new AmpImage(800, 400);
            img.Src = "https://placekitten.com/800/400";
            img.Alt = "Test alt";
            img.Attributes.Layout = LayoutType.Responsive;
            car.Images.Add(img);
            img = new AmpImage(800, 400);
            img.Src = "https://amp.dev/static/img/docs/tutorials/firstemail/photo_by_craig_mclaclan.jpg";
            img.Alt = "Test 3 alt";
            img.Attributes.Layout = LayoutType.Fill;
            car.Images.Add(img);
            msg.AddAmpComponent(car);

            //add AmpFitText component
            AmpFitText txt = new AmpFitText("Lorem ipsum dolor sit amet, has nisl nihil convenire et, vim at aeque inermis reprehendunt.");
            txt.Attributes.Width = 600;
            txt.Attributes.Height = 300;
            txt.Attributes.Layout = LayoutType.Responsive;
            txt.MinFontSize = 8;
            txt.MaxFontSize = 16;
            txt.Value = "Lorem ipsum dolor sit amet, has nisl nihil convenire et, vim at aeque inermis reprehendunt.";
            msg.AddAmpComponent(txt);

            //add AmpAccordion component
            AmpAccordion acc = new AmpAccordion();
            acc.ExpandSingleSection = true;

            Section sec = new Section();
            sec.Header = new SectionHeader(SectionHeaderType.h2, "Section 1");
            sec.Value = new SectionValue("Content in section 1.");
            acc.Sections.Add(sec);

            sec = new Section();
            sec.Header = new SectionHeader(SectionHeaderType.h2, "Section 2");
            sec.Value = new SectionValue("Content in section 2.");
            acc.Sections.Add(sec);

            img = new AmpImage(800, 400);
            img.Src = "https://placekitten.com/800/400";
            img.Alt = "Test alt";
            img.Attributes.Layout = LayoutType.Responsive;

            sec = new Section();
            sec.Header = new SectionHeader(SectionHeaderType.h2, "Section 3");
            sec.Value = new SectionValue(img);
            acc.Sections.Add(sec);
            msg.AddAmpComponent(acc);

            //add AmpForm component
            AmpForm form = new AmpForm();

            form.Method = FormMethod.Post;
            form.ActionXhr = "https://example.com/subscribe";
            form.Target = FormTarget.Top;

            FormField field = new FormField("Name:", "text");
            field.Name = "name";
            field.IsRequired = true;
            form.Fieldset.Add(field);

            field = new FormField("Email:", "email");
            field.Name = "email";
            field.IsRequired = true;
            form.Fieldset.Add(field);

            field = new FormField();
            field.InputType = "submit";
            field.Value = "Subscribe";
            form.Fieldset.Add(field);
            msg.AddAmpComponent(form);

            msg.Save(dataDir + "AmpTest_1.eml");
            MailMessage savedmsg = MailMessage.Load(dataDir + "AmpTest_1.eml");
            AmpMessage ampMsg = savedmsg as AmpMessage;
            if (ampMsg != null)
            {
                DateTime dt = new DateTime(2019, 9, 27, 1, 1, 1, DateTimeKind.Utc);
                AmpTimeago time = new AmpTimeago(dt);
                time.Attributes.Width = 600;
                time.Attributes.Height = 300;
                time.Attributes.Layout = LayoutType.Fixed;
                time.Locale = "en";
                time.Cutoff = 600;
                ampMsg.AddAmpComponent(time);

                ampMsg.Save(dataDir + "AmpTest_2.eml");
            }
            // ExEnd:1
        }
    }
}
