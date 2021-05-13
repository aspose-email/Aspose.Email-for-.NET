using Aspose.Email.Live.Demos.UI.Helpers;
using Aspose.Email.Mapi;
using System;
using System.IO;

namespace Aspose.Email.Live.Demos.UI.Services.Email
{
    public partial class EmailService
    {
		public void RemoveAnnotation(Stream mailStream, string inputFileNameWithExtension, IOutputHandler handler)
		{
			using (var mail = MapiHelper.GetMapiMessageFromStream(mailStream))
			{
				var options = FollowUpManager.GetOptions(mail);

				var flag = options.FlagRequest;

				if (flag != null)
				{
					using (var output = handler.CreateOutputStream(Path.ChangeExtension(inputFileNameWithExtension, ".txt")))
						using (var writer = new StreamWriter(output))
							writer.WriteLine($"{options.StartDate}{Environment.NewLine}{flag}");

					FollowUpManager.ClearFlag(mail);
				}

				mailStream.Position = 0;

				using (var output = handler.CreateOutputStream(inputFileNameWithExtension))
					mailStream.CopyTo(output);
			}
		}
	}
}
