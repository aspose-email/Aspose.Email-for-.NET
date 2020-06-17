using Aspose.Email.Live.Demos.UI.FileProcessing;
using Aspose.Email.Live.Demos.UI.LibraryHelpers;
using Aspose.Email.Live.Demos.UI.Models;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Mbox;
using Aspose.Email.Storage.Pst;
using Aspose.Html;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Http;
using Aspose.Email.Live.Demos.UI.Models.Common;

namespace Aspose.Email.Live.Demos.UI.Models
{
	///<Summary>
	/// AsposeEmailMerger class to merge email file
	///</Summary>
	public class AsposeEmailMerger : EmailBase
	{
		static string[] AllowedTypes =
		{
			".msg",
			".ost",
			".pst",
			".eml",
			".mbox"
		};

		///<Summary>
		/// initialize AsposeEmailMerger class 
		///</Summary>
		public AsposeEmailMerger()
		{
			Aspose.Email.Live.Demos.UI.Models.License.SetAsposeEmailLicense();
		}

		
		public Response Merge(InputFiles inputFiles, string outputType)
		{

			var processor = new CustomSingleOrZipFileProcessor()
			{
				
				CustomFilesProcessMethod = (string[] inputFilePaths, string outputFolderPath) =>
				{
					if (inputFilePaths.Length < 2)
						throw new Exception("Required minimum two files for merging");

					if (inputFilePaths.Length > 10)
						throw new Exception("10 files is maximum for merging. Please, remove excess files");

					var groupsDictionary = inputFilePaths.GroupBy(x => Path.GetExtension(x)?.ToLowerInvariant() ?? "null").Select(ShrinkKeys).ToDictionary(x => x.Item1, x => x.Item2);

					if (groupsDictionary.ContainsKey("null"))
						throw new Exception("All files must have extension in name");

					Aspose.Email.Live.Demos.UI.Models.License.SetAsposeEmailLicense();

					foreach (var item in groupsDictionary)
					{
						MergeFiles(item.Key, item.Value, outputFolderPath);
					}
				}
			};

			(string folder, string[] files) = (null, null);
			folder = inputFiles.FolderName;
			files = inputFiles.FileName;
			try
			{
				
				return processor.Process(folder, files);
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
				return new Response() { Status = "200", Text = "Error on processing file" };
			}
		}

		(string, string[]) ShrinkKeys(IGrouping<string, string> pair)
		{
			switch (pair.Key.ToLowerInvariant())
			{
				case ".eml":
				case ".msg":
					return (".mail", pair.ToArray());
				case ".ost":
				case ".pst":
				case ".mbox":
					return (".storage", pair.ToArray());
				default:
					throw new Exception("Wrong file type " + pair.Key);
			}
		}

		void MergeFiles(string extension, string[] values, string outputFolderPath)
		{
			switch (extension)
			{
				case ".mail":
					MergeMails(values, outputFolderPath);
					break;
				case ".storage":
					MergeStorages(values, outputFolderPath);
					break;
				default:
					break;
			}
		}

		void MergeMails(string[] files, string outputFolderPath)
		{
			if (files.Length < 1)
				return;

			var mails = files.Select(x =>
			{
				var m = MapiHelper.GetMapiMessageFromFile(x);
				if (m == null)
					throw new Exception("Invalid file format " + Path.GetFileName(x));
				return m;
			}).ToArray();

			var folderPath = Path.Combine(Config.Configuration.OutputDirectory, Guid.NewGuid().ToString());
			var filePath = Path.Combine(folderPath, "Merged.html");

			Directory.CreateDirectory(folderPath);

			var mail = mails[0];

			MergeHtmlContentsAndSaveInFile(mails.Select(x => x.BodyHtml), filePath);
			mail.SetBodyContent(System.IO.File.ReadAllText(filePath), BodyContentType.Html);

			Directory.Delete(folderPath, true);

			for (int i = 1; i < mails.Length; i++)
			{
				var addedMail = mails[i];

				if (addedMail.Attachments != null)
					foreach (var attachment in addedMail.Attachments)
						mail.Attachments.Add(attachment);
			}

			mail.Save(Path.Combine(outputFolderPath, "MergedMessage.msg"));
		}

		void MergeStorages(string[] files, string outputFolderPath)
		{
			if (files.Length < 1)
				return;

			var extension = Path.GetExtension(files[0]);

			if (extension.ToLowerInvariant() == ".mbox")
			{
				using (var writer = new MboxrdStorageWriter(Path.Combine(outputFolderPath, "MergedStorage.mbox"), false))
				{
					for (int i = 0; i < files.Length; i++)
						AppendFile(writer, files[i]);
				}
			}
			else
			{
				using (var storage = PersonalStorage.Create(Path.Combine(outputFolderPath, "MergedStorage.pst"), FileFormatVersion.Unicode))
				{
					for (int i = 0; i < files.Length; i++)
						AppendFile(storage, files[i]);
				}
			}
		}




		///<Summary>
		/// MergeHtmlContentsAndSaveInFile method to merge html contents and save in file
		///</Summary>
		public static void MergeHtmlContentsAndSaveInFile(IEnumerable<string> htmls, string temporaryHtmlPath)
		{
			var docs = htmls.Select(x => new Aspose.Html.HTMLDocument(x, "")).ToArray();

			var htmlDocument1 = docs[0];

			for (int i = 1; i < docs.Length; i++)
				AppendDoc(htmlDocument1, docs[i]);

			htmlDocument1.Save(temporaryHtmlPath);
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

		private void AppendFile(PersonalStorage storage, string filePath)
		{
			var extension = Path.GetExtension(filePath);

			if (extension.ToLowerInvariant() == ".mbox")
			{
				using (var reader = new MboxrdStorageReader(filePath, false))
				{
					for (int i = 0; i < reader.GetTotalItemsCount(); i++)
						storage.RootFolder.AddMessage(MapiMessage.FromMailMessage(reader.ReadNextMessage()));
				}
			}
			else
			{
				storage.MergeWith(new[] { filePath });
			}
		}

		private void AppendFile(MboxrdStorageWriter writer, string filePath)
		{
			var extension = Path.GetExtension(filePath);

			if (extension.ToLowerInvariant() == ".mbox")
			{
				using (var reader = new MboxrdStorageReader(filePath, false))
				{
					for (int i = 0; i < reader.GetTotalItemsCount(); i++)
						writer.WriteMessage(reader.ReadNextMessage());
				}
			}
			else
			{
				using (var storage = PersonalStorage.FromFile(filePath))
				{
					var options = new MailConversionOptions();
					HandleFolderAndSubfolders(mapiMessage =>
					{
						var msg = mapiMessage.ToMailMessage(options);
						writer.WriteMessage(msg);
					}, storage.RootFolder, options);
				}
			}
		}

		void HandleFolderAndSubfolders(Action<MapiMessage> handler, FolderInfo folderInfo, MailConversionOptions options)
		{
			foreach (MapiMessage mapiMessage in folderInfo.EnumerateMapiMessages())
			{
				handler(mapiMessage);
			}

			if (folderInfo.HasSubFolders == true)
			{
				foreach (FolderInfo subfolderInfo in folderInfo.GetSubFolders())
				{
					HandleFolderAndSubfolders(handler, subfolderInfo, options);
				}
			}
		}
	}
}


