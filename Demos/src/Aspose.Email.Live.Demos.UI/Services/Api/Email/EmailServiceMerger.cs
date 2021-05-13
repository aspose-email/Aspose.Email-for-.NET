using Aspose.Email.Live.Demos.UI.Helpers;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Mbox;
using Aspose.Email.Storage.Pst;
using Aspose.Html;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;


namespace Aspose.Email.Live.Demos.UI.Services.Email
{
	public partial class EmailService
	{
		public void Merge(IEnumerable<(Stream input, string inputFileNameWithExtension)> inputs, IOutputHandler handler)
		{
			foreach (var item in inputs.GroupBy(GetKey))
				MergeFiles(item.Key, item.ToArray(), handler);
		}

		string GetKey((Stream input, string inputFileNameWithExtension) x)
		{
			var extension = Path.GetExtension(x.inputFileNameWithExtension);

			switch (extension)
			{
				case ".eml":
				case ".msg":
					return ".mail";
				case ".ost":
				case ".pst":
				case ".mbox":
					return ".storage";
				default:
					throw new Exception("All files must have extension in name");
			}
		}

		void MergeFiles(string extension, IEnumerable<(Stream input, string inputFileNameWithExtension)> inputs, IOutputHandler handler)
		{
			switch (extension)
			{
				case ".mail":
					MergeMails(inputs, handler);
					break;
				case ".storage":
					MergeStorages(inputs, handler);
					break;
				default:
					break;
			}
		}

		void MergeMails(IEnumerable<(Stream input, string inputFileNameWithExtension)> inputs, IOutputHandler handler)
		{
			if (inputs.IsNullOrEmpty())
				return;

			var mails = inputs.Select(x => MapiHelper.GetMapiMessageFromStream(x.input));

			var mail = mails.First();

			var mergedContent = MergeHtmlContentsAndSave(mails.Select(x => x.BodyHtml));

			mail.SetBodyContent(mergedContent, BodyContentType.Html);

			foreach (var addedMail in mails.Skip(1))
			{
				if (addedMail.Attachments != null)
					foreach (var attachment in addedMail.Attachments)
						mail.Attachments.Add(attachment);
			}

			using (var output = handler.CreateOutputStream("MergedMessage.msg"))
				mail.Save(output);
		}

		void MergeStorages(IEnumerable<(Stream input, string inputFileNameWithExtension)> inputs, IOutputHandler handler)
		{
			if (inputs.IsNullOrEmpty())
				return;

			var extension = Path.GetExtension(inputs.First().inputFileNameWithExtension);

			if (".mbox".Equals(extension, StringComparison.OrdinalIgnoreCase))
			{
				using (var output = handler.CreateOutputStream("MergedStorage.mbox"))
				{
					using (var writer = new MboxrdStorageWriter(output, false))
					{
						foreach (var item in inputs)
							AppendFile(writer, item.input, item.inputFileNameWithExtension);
					}
				}
			}
			else
			{
				using (var output = handler.CreateOutputStream("MergedStorage.pst"))
				{
					using (var storage = PersonalStorage.Create(output, FileFormatVersion.Unicode))
					{
						foreach (var item in inputs)
							AppendFile(storage, item.input, item.inputFileNameWithExtension);
					}
				}
			}
		}

		///<Summary>
		/// MergeHtmlContentsAndSaveInFile method to merge html contents and save in file
		///</Summary>
		public string MergeHtmlContentsAndSave(IEnumerable<string> htmls)
		{
			var docs = htmls.Select(x => new HTMLDocument(x, ""));

			var mergedDoc = docs.Aggregate((x, y) =>
			{
				AppendDoc(x, y);
				return x;
			});

			return mergedDoc.ToHTMLContentString();
		}

		private static void AppendDoc(HTMLDocument htmlDocument1, HTMLDocument htmlDocument2)
		{
			foreach (var item in htmlDocument2.Body.Children.ToArray())
				htmlDocument1.Body.AppendChild(item);

			if (htmlDocument1.Title != htmlDocument2.Title)
				htmlDocument1.Title += htmlDocument2.Title;

			var headElement = htmlDocument1.GetElementsByTagName("head")[0];
			var metaCollection1 = htmlDocument2.QuerySelectorAll("meta[name]").Cast<Html.HTMLMetaElement>();
			var metaCollection2 = htmlDocument2.QuerySelectorAll("meta[name]").Cast<Html.HTMLMetaElement>();

			foreach (var metaElement1 in metaCollection1)
			{
				Html.HTMLMetaElement metaElement2;
				try
				{
					metaElement2 = metaCollection2.First(e => e.Name == metaElement1.Name);
					metaElement1.Content += metaElement2.Content;
				}
				catch (System.InvalidOperationException)
				{
					headElement.AppendChild(metaElement1);
				}
			}
		}

		private void AppendFile(PersonalStorage storage, Stream input, string inputFileNameWithExtension)
		{
			var extension = Path.GetExtension(inputFileNameWithExtension);

			if (extension.ToLowerInvariant() == ".mbox")
			{
				using (var reader = new MboxrdStorageReader(input, false))
				{
					for (int i = 0; i < reader.GetTotalItemsCount(); i++)
						storage.RootFolder.AddMessage(MapiMessage.FromMailMessage(reader.ReadNextMessage()));
				}
			}
			else
			{
				storage.MergeWith(new[] { input });
			}
		}

		private void AppendFile(MboxrdStorageWriter writer, Stream input, string inputFileNameWithExtension)
		{
			var extension = Path.GetExtension(inputFileNameWithExtension);

			if (extension.ToLowerInvariant() == ".mbox")
			{
				using (var reader = new MboxrdStorageReader(input, false))
				{
					for (int i = 0; i < reader.GetTotalItemsCount(); i++)
						writer.WriteMessage(reader.ReadNextMessage());
				}
			}
			else
			{
				using (var storage = PersonalStorage.FromStream(input))
				{
					var options = new MailConversionOptions();
					HandleFolderAndSubfolders(mapiMessage =>
					{
						var msg = mapiMessage.ToMailMessage(options);
						writer.WriteMessage(msg);
					}, storage.RootFolder);
				}
			}
		}
	}
}
