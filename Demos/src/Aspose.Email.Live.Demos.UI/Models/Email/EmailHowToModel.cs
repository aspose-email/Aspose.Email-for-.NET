using System;
using System.Collections.Generic;

namespace Aspose.Email.Live.Demos.UI.Models.Email
{
    public class EmailHowToModel
    {
        public string Title { get; set; }
        public List<HowToItem> List { get; set; }

        public ViewModel Parent;
        public string AppName => Parent.AppName;

        public EmailHowToModel(ViewModel parent)
        {
            Parent = parent;
            Title = parent.AppResources.GetResourceOrDefault("HowToTitle");

            List = new List<HowToItem>();

            for (int i = 1; ; i++)
            {
                var text = parent.AppResources.GetResourceOrDefault("HowToFeature" + i);

                if (text.IsNullOrWhiteSpace())
                    break;

                var item = new HowToItem();

                item.Text = text;
                List.Add(item);
            }
        }
    }

    public class HowToItem
    {
        public string Name { get; set; }
        public string Text { get; set; }
        public string ImageUrl { get; set; }
        public string UrlPath { get; set; }
    }
}
