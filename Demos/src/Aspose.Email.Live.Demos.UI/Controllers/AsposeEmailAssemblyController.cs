using Aspose.Email.Live.Demos.UI.FileProcessing;
using Aspose.Email.Live.Demos.UI.LibraryHelpers;
using Aspose.Email.Live.Demos.UI.Models;
using Aspose.Email;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Mbox;
using Aspose.Email.Storage.Pst;
using Aspose.Email.Tools.Merging;
using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace Aspose.Email.Live.Demos.UI.Controllers
{
	public class AsposeEmailAssemblyController : EmailBase
	{
		public AsposeEmailAssemblyController()
		{
			Aspose.Email.Live.Demos.UI.Models.License.SetAsposeEmailLicense();
		}

		public async Task<Response> Assemble(string sourceName, int tableIndex = 0, string delimiter = ",")
		{
			(string folder, string[] files) = (null, null);

			try
			{
				(folder, files) = await UploadFiles();

				if (files.Length < 2)
					return new Response
					{
						FileName = null,
						FolderName = folder,
						Status = "Request doesn't contain template or dataSource",
						StatusCode = 400
					};

				var templatePath = files[0];
				var dataSourcePath = files[1];

				var dataTable = AssemblyDataHelper.PrepareDataTable(dataSourcePath, sourceName, tableIndex, delimiter);

				if (dataTable == null)
					return new Response
					{
						FileName = null,
						FolderName = folder,
						Status = "Can't process the data source",
						StatusCode = 500
					};

				var processor = new CustomSingleOrZipFileProcessor()
				{
					
					CustomFileProcessMethod = (string inputFilePath, string outputFolderPath) =>
					{
						var ext = Path.GetExtension(inputFilePath);

						Aspose.Email.Live.Demos.UI.Models.License.SetAsposeEmailLicense();

						switch (ext)
						{
							case ".eml":
							case ".msg":
								AssemblyMailFile(dataTable, inputFilePath, outputFolderPath);
								break;
							case ".pst":
							case ".ost":
								AssemblyStorageFile(ext, dataTable, inputFilePath, outputFolderPath);
								break;
							case ".mbox":
								AssemblyMboxFile(dataTable, inputFilePath, outputFolderPath);
								break;
							default:
								break;
						}
					}
				};


				return processor.Process(folder, templatePath);
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);	
				return new Response() { StatusCode = 500, Status = "Error on processing file", Text = "Error on processing file" };
			}
		}

		void AssemblyMboxFile(DataTable dataTable, string inputFilePath, string outputFolderPath)
		{
			for (int i = 0; i < dataTable.Rows.Count; i++)
			{
				using (var writer = new MboxrdStorageWriter(Path.Combine(outputFolderPath, Path.GetFileNameWithoutExtension(inputFilePath) + " Assembled" + i + ".mbox"), false))
				{
					using (var reader = new MboxrdStorageReader(inputFilePath, false))
					{
						for (int m = 0; m < reader.GetTotalItemsCount(); m++)
						{
							var message = reader.ReadNextMessage();
							var engine = new TemplateEngine(message);
							var messages = engine.Instantiate(dataTable);
							var mes = messages.ElementAtOrDefault(i);

							if (mes != null)
								writer.WriteMessage(mes);
						}
					}
				}
			}
		}

		void AssemblyStorageFile(string extension, DataTable dataTable, string inputFilePath, string outputFolderPath)
		{
			var conversionOptions = new MailConversionOptions();

			for (int i = 0; i < dataTable.Rows.Count; i++)
			{
				using (var personalStorage = PersonalStorage.FromFile(inputFilePath))
				{
					HandleFolderAndSubfolders(personalStorage, (mapiMessage, entryId, folder) =>
					{
						var engine = new TemplateEngine(mapiMessage.ToMailMessage(conversionOptions));
						var messages = engine.Instantiate(dataTable);

						var mes = messages.ElementAtOrDefault(i);

						if (mes != null)
							folder.UpdateMessage(entryId, MapiMessage.FromMailMessage(mes));
					}, personalStorage.RootFolder, new MailConversionOptions());

					if (extension == ".pst")
						personalStorage.SaveAs(Path.Combine(outputFolderPath, Path.GetFileNameWithoutExtension(inputFilePath) + " Assembled" + i + ".pst"), FileFormat.Pst);
					else
						personalStorage.SaveAs(Path.Combine(outputFolderPath, Path.GetFileNameWithoutExtension(inputFilePath) + " Assembled" + i + ".ost"), FileFormat.Ost);
				}
			}
		}

		void HandleFolderAndSubfolders(PersonalStorage storage, Action<MapiMessage, string, FolderInfo> handler, FolderInfo folderInfo, MailConversionOptions options)
		{
			foreach (var entryId in folderInfo.EnumerateMessagesEntryId())
			{
				var mapiMessage = storage.ExtractMessage(entryId);
				handler(mapiMessage, entryId, folderInfo);
			}

			if (folderInfo.HasSubFolders == true)
			{
				foreach (FolderInfo subfolderInfo in folderInfo.GetSubFolders())
				{
					HandleFolderAndSubfolders(storage, handler, subfolderInfo, options);
				}
			}
		}

		private void AssemblyMailFile(DataTable dataTable, string inputFilePath, string outputFolderPath)
		{
			var mail = MapiHelper.GetMapiMessageFromFile(inputFilePath);

			if (mail == null)
				throw new Exception("Invalid file format");

			var engine = new TemplateEngine(inputFilePath);
			var messages = engine.Instantiate(dataTable);

			int i = 0;

			var resultFileName = Path.GetFileNameWithoutExtension(inputFilePath) + " Assembled";

			foreach (var assembledMessage in messages)
				assembledMessage.Save(Path.Combine(outputFolderPath, resultFileName + i + ".msg"), MsgSaveOptions.DefaultMsgUnicode);
		}

		public Task<Response> Assemble(string folderName, string templateFilename, string datasourceFilename, string datasourceName, int datasourceTableIndex = 0, string delimiter = ",")
		{
			var dataTable = AssemblyDataHelper.PrepareDataTable(Path.Combine(Config.Configuration.WorkingDirectory, folderName, datasourceFilename), datasourceName, datasourceTableIndex, delimiter);

			if (dataTable == null)
				return Task.FromResult(new Response
				{
					FileName = null,
					FolderName = folderName,
					Status = "Can't process the data source",
					StatusCode = 500
				});

			var processor = new CustomSingleOrZipFileProcessor()
			{
				
				CustomFileProcessMethod = (string inputFilePath, string outputFolderPath) =>
				{
					var mail = MapiHelper.GetMapiMessageFromFile(inputFilePath);

					if (mail == null)
						throw new Exception("Invalid file format");

					Aspose.Email.Live.Demos.UI.Models.License.SetAsposeEmailLicense();

					var engine = new TemplateEngine(inputFilePath);
					var messages = engine.Instantiate(dataTable);

					int i = 0;

					var resultFileName = Path.GetFileNameWithoutExtension(templateFilename) + " Assembled";

					foreach (var assembledMessage in messages)
						assembledMessage.Save(Path.Combine(outputFolderPath, resultFileName + i+".msg"), MsgSaveOptions.DefaultMsgUnicode);
				}
			};

			return Task.FromResult(processor.Process(folderName, templateFilename));
		}
	}
}
