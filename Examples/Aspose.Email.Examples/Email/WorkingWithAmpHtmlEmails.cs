// Demonstrates how to compose an AMP HTML email message with interactive components
// (animation, image, carousel, fit text, accordion, form, and timeago),
// save it, reload it, and append an additional component.

using System;
using Aspose.Email.Amp;

namespace Aspose.Email.Examples.Email
{
    internal static class WorkingWithAmpHtmlEmails
    {
        public static void Run()
        {
            var eml = new AmpMessage();
            eml.HtmlBody = "<html><body> Hello AMP </body></html>";

            // AmpAnim — animated GIF component
            var anim = new AmpAnim(800, 400)
            {
                Src = "https://placekitten.com/800/400",
                Alt = "Test alt",
                Attribution = "The Go gopher was designed by Reneee French",
                Fallback = "offline"
            };
            anim.Attributes.Layout = LayoutType.Responsive;
            eml.AddAmpComponent(anim);

            // AmpImage — static image component
            var img = new AmpImage(800, 400)
            {
                Src = "https://placekitten.com/800/400",
                Alt = "Test alt"
            };
            img.Attributes.Layout = LayoutType.Responsive;
            eml.AddAmpComponent(img);

            // AmpCarousel — image carousel component
            var car = new AmpCarousel(800, 400);

            var img1 = new AmpImage(800, 400)
            {
                Src = "https://amp.dev/static/img/docs/tutorials/firstemail/photo_by_caleb_woods.jpg",
                Alt = "Photo by Caleb Woods"
            };
            img1.Attributes.Layout = LayoutType.Fixed;
            car.Images.Add(img1);

            var img2 = new AmpImage(800, 400)
            {
                Src = "https://placekitten.com/800/400",
                Alt = "Kitten"
            };
            img2.Attributes.Layout = LayoutType.Responsive;
            car.Images.Add(img2);

            var img3 = new AmpImage(800, 400)
            {
                Src = "https://amp.dev/static/img/docs/tutorials/firstemail/photo_by_craig_mclaclan.jpg",
                Alt = "Photo by Craig McLaclan"
            };
            img3.Attributes.Layout = LayoutType.Fill;
            car.Images.Add(img3);
            eml.AddAmpComponent(car);

            // AmpFitText — auto-scaling text component
            var txt = new AmpFitText("Lorem ipsum dolor sit amet, has nisl nihil convenire et, vim at aeque inermis reprehendunt.")
            {
                MinFontSize = 8,
                MaxFontSize = 16
            };
            txt.Attributes.Width = 600;
            txt.Attributes.Height = 300;
            txt.Attributes.Layout = LayoutType.Responsive;
            eml.AddAmpComponent(txt);

            // AmpAccordion — collapsible sections component
            var acc = new AmpAccordion { ExpandSingleSection = true };

            acc.Sections.Add(new Section
            {
                Header = new SectionHeader(SectionHeaderType.h2, "Section 1"),
                Value = new SectionValue("Content in section 1.")
            });
            acc.Sections.Add(new Section
            {
                Header = new SectionHeader(SectionHeaderType.h2, "Section 2"),
                Value = new SectionValue("Content in section 2.")
            });

            var secImg = new AmpImage(800, 400)
            {
                Src = "https://placekitten.com/800/400",
                Alt = "Section image"
            };
            secImg.Attributes.Layout = LayoutType.Responsive;
            acc.Sections.Add(new Section
            {
                Header = new SectionHeader(SectionHeaderType.h2, "Section 3"),
                Value = new SectionValue(secImg)
            });
            eml.AddAmpComponent(acc);

            // AmpForm — subscription form component
            var form = new AmpForm
            {
                Method = FormMethod.Post,
                ActionXhr = "https://example.com/subscribe",
                Target = FormTarget.Top
            };
            form.Fieldset.Add(new FormField("Name:", "text") { Name = "name", IsRequired = true });
            form.Fieldset.Add(new FormField("Email:", "email") { Name = "email", IsRequired = true });
            form.Fieldset.Add(new FormField { InputType = "submit", Value = "Subscribe" });
            eml.AddAmpComponent(form);

            var firstPath = Data.Out/"AmpTest_1.eml";
            eml.Save(firstPath);
            Console.WriteLine($"Saved AMP message with animation, image, carousel, fit text, accordion and form:");
            Console.WriteLine($"  {firstPath}");

            // Loading an AMP message back gives an AmpMessage, so more components can be
            // added to a message that already exists.
            if (MailMessage.Load(firstPath) is AmpMessage savedEml)
            {
                var time = new AmpTimeago(new DateTime(2019, 9, 27, 1, 1, 1, DateTimeKind.Utc))
                {
                    Locale = "en",
                    Cutoff = 600
                };
                time.Attributes.Width = 600;
                time.Attributes.Height = 300;
                time.Attributes.Layout = LayoutType.Fixed;
                savedEml.AddAmpComponent(time);

                var secondPath = Data.Out/"AmpTest_2.eml";
                savedEml.Save(secondPath);

                Console.WriteLine("Reloaded it and appended a timeago component:");
                Console.WriteLine($"  {secondPath}");
            }
        }
    }
}
