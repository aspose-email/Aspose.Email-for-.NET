using Aspose.Email.Live.Demos.UI.Helpers;
using Aspose.Email;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Mbox;
using Aspose.Email.Storage.Pst;
using Aspose.Email.Tools.Merging;
using System;
using System.IO;
using System.Linq;
using System.Data;

namespace Aspose.Email.Live.Demos.UI.Services.Email
{
	public partial class EmailService
    {
		public void AssembleFile(Stream input, string shortResourceNameWithExtension, IOutputHandler handler, DataTable dataTable)
		{
			var ext = Path.GetExtension(shortResourceNameWithExtension);

			switch (ext)
			{
				case ".eml":
				case ".msg":
					AssemblyMailFile(input, shortResourceNameWithExtension, handler, dataTable);
					break;
				case ".pst":
				case ".ost":
					AssemblyStorageFile(input, shortResourceNameWithExtension, handler, dataTable);
					break;
				case ".mbox":
					AssemblyMboxFile(input, shortResourceNameWithExtension, handler, dataTable);
					break;
				default:
					break;
			}
		}

		public void AssemblyStorageFile(Stream input, string shortResourceNameWithExtension, IOutputHandler handler, DataTable dataTable)
		{
			var conversionOptions = new MailConversionOptions();

			for (int i = 0; i < dataTable.Rows.Count; i++)
			{
				using (var output = handler.CreateOutputStream(Path.GetFileNameWithoutExtension(shortResourceNameWithExtension) + " Assembled" + i + ".pst"))
				{
					input.CopyTo(output);

					using (var personalStorage = PersonalStorage.FromStream(output))
					{
						HandleFolderAndSubfolders(personalStorage, (mapiMessage, entryId, folder) =>
						{
							var engine = new TemplateEngine(mapiMessage.ToMailMessage(conversionOptions));
							var messages = engine.Instantiate(dataTable);

							var mes = messages.ElementAtOrDefault(i);

							if (mes != null)
								folder.UpdateMessage(entryId, MapiMessage.FromMailMessage(mes));
						}, personalStorage.RootFolder, new MailConversionOptions());
					}
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

		public void AssemblyMailFile(Stream input, string shortResourceNameWithExtension, IOutputHandler handler, DataTable dataTable)
		{
			using (var mail = MapiHelper.GetMailMessageFromStream(input))
			{
				var engine = new TemplateEngine(mail);
				var messages = engine.Instantiate(dataTable);

				int i = 0;

				var resultFileName = Path.GetFileNameWithoutExtension(shortResourceNameWithExtension) + " Assembled";

				foreach (var assembledMessage in messages)
					using (var output = handler.CreateOutputStream(resultFileName + ++i + ".msg"))
						assembledMessage.Save(output, MsgSaveOptions.DefaultMsgUnicode);
			}
		}

		public void AssemblyMboxFile(Stream input, string shortResourceNameWithExtension, IOutputHandler handler, DataTable dataTable)
		{
			for (int i = 0; i < dataTable.Rows.Count; i++)
			{
				using (var output = handler.CreateOutputStream(Path.GetFileNameWithoutExtension(shortResourceNameWithExtension) + " Assembled" + i + ".mbox"))
				{
					using (var writer = new MboxrdStorageWriter(output, false))
					{
						using (var reader = new MboxrdStorageReader(input, new MboxLoadOptions() { LeaveOpen = false }))
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
		}
	}
}
