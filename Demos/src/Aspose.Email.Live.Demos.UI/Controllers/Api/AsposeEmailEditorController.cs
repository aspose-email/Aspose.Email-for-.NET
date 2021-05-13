using Aspose.Email.Live.Demos.UI.Helpers;
using Aspose.Email.Live.Demos.UI.Models;
using Aspose.Email.Live.Demos.UI.Services;
using Aspose.Email.Live.Demos.UI.Models;
using Aspose.Email.Live.Demos.UI.Services.Email;
using Aspose.Email;
using Aspose.Email.Mapi;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using Aspose.Email.Live.Demos.UI;

namespace Aspose.Email.Live.Demos.UI.Controllers
{
    public class AsposeEmailEditorController : AsposeEmailBaseApiController
	{
        public AsposeEmailEditorController(ILogger<AsposeEmailEditorController> logger,
			IConfiguration config,
			IStorageService storageService,
			IEmailService emailService)
			: base(logger, config, storageService, emailService)
		{
        }

        [HttpPost("GetDocumentData")]
        public async Task<string> GetDocumentData(string fileName, string folderName)
		{
			fileName = HttpUtility.HtmlDecode(fileName);

			if (fileName.IsNullOrWhiteSpace())
				throw new BadRequestException("File name not provided");

			if (folderName.IsNullOrWhiteSpace())
				throw new BadRequestException("Folder name not provided");

			using (var fileStream = await StorageService.ReadFile(folderName, fileName))
            {
				return JsonConvert.SerializeObject(EmailService.ParseFileData(fileStream));
			}
		}

		[HttpPost("UpdateContentsWithAttachments")]
		public async Task<string> UpdateContentsWithAttachments([FromBody] EditorUpdateContentRequest data)
		{
			if (!data.CreateNewFile)
			{
				if (data.FileName.IsNullOrWhiteSpace())
					throw new BadRequestException("File name not provided");

				if (data.FolderName.IsNullOrWhiteSpace())
					throw new BadRequestException("Folder name not provided");
			}

			if (data.Htmldata == null)
				throw new BadRequestException("Html data name not provided");

			if (data.OutputType.IsNullOrWhiteSpace())
				throw new BadRequestException("Output type name not provided");

			if (!data.CreateNewFile)
				data.FileName = HttpUtility.HtmlDecode(data.FileName);
			else
				data.FileName = "New file";

			data.OutputType = data.OutputType.ToLower();
			var savingFileName = Path.GetFileNameWithoutExtension(data.FileName) + data.OutputType;
			var uid = Guid.NewGuid().ToString();

			try
			{
				switch (data.OutputType)
				{
					case ".html":
						await StorageService.SaveFile(data.Htmldata, uid, savingFileName);
						break;
					default:
						{
							using (var message = await GetMessageFromData(data))
							{
								message.SetBodyContent(data.Htmldata, BodyContentType.Html);

								var changes = data.AttachmentsData?.Select(x => x.Value).ToArray();
								var attachments = message.Attachments.ToArray();

								if (changes != null)
								{
									for (int i = 0; i < changes.Length; i++)
									{
										var ch = changes[i];

										if (ch.NeedRemove)
											message.Attachments.Remove(attachments[ch.Index]);

										if (ch.WasAdded && ch.Name != null && ch.SourceBase64 != null)
											message.Attachments.Add(ch.Name, Convert.FromBase64String(ch.SourceBase64));
									}
								}

								await StorageService.SaveMessage(message, uid, savingFileName, data.OutputType == ".eml" ? EmlSaveOptions.DefaultEml : EmlSaveOptions.DefaultMsgUnicode);
								break;
							}
						}
				}
			
				return JsonConvert.SerializeObject(new Response()
				{
					FileName = HttpUtility.UrlEncode(savingFileName),
					FolderName = uid,
					StatusCode = 200
				});

			}
			catch (Exception ex)
			{
				var logMsg = $"ControllerName = {nameof(AsposeEmailEditorController)}, MethodName = {nameof(UpdateContents)}, Folder = {uid}";
				Logger.LogError(ex, logMsg, AsposeEmail + EditorApp, savingFileName);

				return JsonConvert.SerializeObject(new Response()
				{
					FileName = HttpUtility.UrlEncode(savingFileName),
					FolderName = uid,
					StatusCode = 500,
					Status = ex.Message
				}.WithTrace(Configuration, ex));
			}
		}

        private async Task<MapiMessage> GetMessageFromData(EditorUpdateContentRequest data)
        {
			if (data.CreateNewFile)
				return new MapiMessage(OutlookMessageFormat.Unicode);

			return await StorageService.ReadMapiMessage(data.FolderName, data.FileName);
		}

        [HttpPost("UpdateContents")]
		public async Task<string> UpdateContents(string fileName, string folderName, string htmldata, string outputType)
        {
			if (fileName.IsNullOrWhiteSpace())
				throw new BadRequestException("File name not provided");

			if (folderName.IsNullOrWhiteSpace())
				throw new BadRequestException("Folder name not provided");

			if (htmldata == null)
				throw new BadRequestException("Html data name not provided");

			if (outputType.IsNullOrWhiteSpace())
				throw new BadRequestException("Output type name not provided");
			
			fileName = HttpUtility.HtmlDecode(fileName);

			outputType = outputType.ToLower();
            var savingFileName = Path.GetFileNameWithoutExtension(fileName) + outputType;
            var uid = Guid.NewGuid().ToString();

            try
            {
                switch (outputType)
                {
                    case ".html":
						await StorageService.SaveFile(htmldata, uid, savingFileName);
						break;
                    default:
                        {
							using (var message = await StorageService.ReadMapiMessage(folderName, fileName))
							{
								message.SetBodyContent(htmldata, BodyContentType.Html);

								await StorageService.SaveMessage(message, uid, savingFileName, outputType == ".eml" ? EmlSaveOptions.DefaultEml : EmlSaveOptions.DefaultMsgUnicode);
								break;
							}
                        }
                }

				return JsonConvert.SerializeObject(new Response()
				{
					FileName = HttpUtility.UrlEncode(savingFileName),
					FolderName = uid,
					StatusCode = 200
				});

            }
			catch(Exception ex)
            {
                var logMsg = $"ControllerName = {nameof(AsposeEmailEditorController)}, MethodName = {nameof(UpdateContents)}, Folder = {uid}";
                Logger.LogError(ex, logMsg, AsposeEmail + EditorApp, savingFileName);

				return JsonConvert.SerializeObject(new Response()
				{
					FileName = HttpUtility.UrlEncode(savingFileName),
					FolderName = uid,
					StatusCode = 500,
					Status = ex.Message
				});
            }
        }
    }
}
