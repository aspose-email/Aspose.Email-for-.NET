using System;
using System.Collections.Generic;

namespace Aspose.Email.Live.Demos.UI.Models.Email
{
    public class EmailFAQPageModel
    {
		public List<FAQItem> List { get; set; }

		public EmailFAQPageModel(ViewModel parent)
			: base()
        {
			List = new List<FAQItem>();

			for (int i = 1; ; i++)
			{
				var question = parent.AppResources.GetResourceOrDefault("FAQQuestion" + i);

				if (question.IsNullOrWhiteSpace())
					break;

				var answer = parent["FAQAnswer" + i];

				List.Add(new FAQItem()
				{
					Question = question,
					AcceptedAnswer = answer
				});
			}

		}
	}

	public class FAQItem
	{
		public string Question { get; set; }
		public string AcceptedAnswer { get; set; }
	}
}
