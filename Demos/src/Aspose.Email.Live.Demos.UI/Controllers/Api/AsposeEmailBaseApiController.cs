using Aspose.Email.Live.Demos.UI.Config;
using Aspose.Email.Live.Demos.UI.FileProcessing;
using Aspose.Email.Live.Demos.UI.Helpers;
using Aspose.Email.Live.Demos.UI.Models;
using Aspose.Email.Live.Demos.UI.Services;
using Aspose.Email.Live.Demos.UI.Services.Email;
using Microsoft.AspNetCore.Http.Features;
using Microsoft.AspNetCore.WebUtilities;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Net.Http.Headers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Email.Live.Demos.UI.Controllers
{

    public abstract class AsposeEmailBaseApiController : ApiControllerBase
	{
		public ILogger Logger { get; }
		public IEmailService EmailService { get; }
		public IStorageService StorageService { get; }
		public IConfiguration Configuration { get; }

		protected AsposeEmailBaseApiController(ILogger logger,
			IConfiguration config,
			IStorageService storageService, 
			IEmailService emailService)
		{
			Logger = logger;
			Configuration = config;
			StorageService = storageService;
			EmailService = emailService;
		}

		private static readonly FormOptions _defaultFormOptions = new FormOptions();

		/// <summary>
		/// Prepare upload files and return as folder and names
		/// </summary>
		public async Task<IDictionary<string, byte[]>> UploadFiles()
		{
			try
			{ 
				if (this.Request.Content.Headers.ContentType == null || !MultipartRequestHelper.IsMultipartContentType(this.Request.Content.Headers.ContentType.MediaType))
					throw new BadRequestException();

				var boundary = MultipartRequestHelper.GetBoundary(MediaTypeHeaderValue.Parse(this.Request.Content.Headers.ContentType.ToString()), _defaultFormOptions.MultipartBoundaryLengthLimit);
				var files = new Dictionary<string, byte[]>();

				using (var stream = await Request.Content.ReadAsStreamAsync())
				{
					var reader = new MultipartReader(boundary, stream);
					var section = await reader.ReadNextSectionAsync();

					while (section != null)
					{
						var hasContentDispositionHeader = ContentDispositionHeaderValue.TryParse(section.ContentDisposition, out var contentDisposition);

						if (hasContentDispositionHeader)
						{
							// This check assumes that there's a file
							// present without form data. If form data
							// is present, this method immediately fails
							// and returns the model error.
							if (!MultipartRequestHelper.HasFileContentDisposition(contentDisposition) && !MultipartRequestHelper.HasFormDataContentDisposition(contentDisposition))
							{
								// ignore section
								//ModelState.AddModelError("File", $"The request couldn't be processed (Error 2).");
								//throw new BadRequestException();
							}
							else
							{
								// Don't trust the file name sent by the client. To display
								// the file name, HTML-encode the value.
								var trustedFileNameForDisplay = WebUtility.HtmlEncode(contentDisposition.FileName.Value ?? contentDisposition.Name.Value);
								//var trustedFileNameForFileStorage = Path.GetRandomFileName();

								// **WARNING!**
								// In the following example, the file is saved without
								// scanning the file's contents. In most production
								// scenarios, an anti-virus/anti-malware scanner API
								// is used on the file before making the file available
								// for download or for use by other systems. 
								// For more information, see the topic that accompanies 
								// this sample.

								var streamedFileContent = await FileHelpers.ProcessStreamedFile(section, ModelState, AppConstants.MaxFileSize);

								if (!ModelState.IsValid)
									throw new BadRequestException();

								var initialName = trustedFileNameForDisplay;
								var name = Path.GetFileNameWithoutExtension(initialName);

								var ext = Path.GetExtension(initialName);
								var counter = 1;

								while (files.ContainsKey(trustedFileNameForDisplay))
								{
									if (string.IsNullOrEmpty(ext))
										trustedFileNameForDisplay = $"{name} ({++counter})";
									else
										trustedFileNameForDisplay = $"{name} ({++counter}).{ext}";
								}

								files[trustedFileNameForDisplay] = streamedFileContent;
							}
						}

						// Drain any remaining section body that hasn't been consumed and
						// read the headers for the next section.
						section = await reader.ReadNextSectionAsync();
					}
				}

				return files;
			}
			catch (Exception ex)
			{
				Logger.LogEmailError(ex, "UploadFiles");
				throw new BadRequestException("Error while uploading the files", ex);
			}
		}

		protected Task<Response> Process(string appName, string folder, string fileName, Action<IEmailService, IOutputHandler, IDictionary<string, byte[]>> process)
		{
			return GetFromStorageAndHandle(appName, folder, fileName, (IDictionary<string, byte[]> files) =>
			{
				var processor = new CustomSingleOrZipFileProcessor(StorageService, Configuration, Logger)
				{
					ProductName = AsposeEmail + appName,
					CustomFilesProcessMethod = (IDictionary<string, byte[]> files, IOutputHandler handler) =>
					{
						process(EmailService, handler, files);
					}
				};

				return processor.Process(files);
			},
			(ex, files) => new Response() { StatusCode = 400, Status = ex.Message },
			(ex, files) => new Response() { StatusCode = 500, Status = "Error on processing: " + ex.Message }.WithTrace(Configuration, ex));
		
		}

		protected Task<Response> Process(string appName, Action<IEmailService, IOutputHandler, IDictionary<string, byte[]>> process)
		{
			return UploadAndHandle(appName, (IDictionary<string, byte[]> files) =>
			{
				var processor = new CustomSingleOrZipFileProcessor(StorageService, Configuration, Logger)
				{
					ProductName = AsposeEmail + appName,
					CustomFilesProcessMethod = (IDictionary<string, byte[]> files, IOutputHandler handler) =>
					{
							process(EmailService, handler, files);
					}
				};

				return processor.Process(files);
			},
			(ex, files) => new Response() { StatusCode = 400, Status = ex.Message },
			(ex, files) => new Response() { StatusCode = 500, Status = "Error on processing: " + ex.Message }.WithTrace(Configuration, ex));
		}

		protected async Task<T> GetFromStorageAndHandle<T>(string appName, string folder, string fileName, 
			Func<IDictionary<string, byte[]>, Task<T>> process,
			Func<BadRequestException, IDictionary<string, byte[]>, T> badRequestHandler, 
			Func<Exception, IDictionary<string, byte[]>, T> exceptionHandler)
		{
			IDictionary<string, byte[]> uploaded = null;

			try
			{
				using (var input = await StorageService.ReadFile(folder, fileName))
				{
					uploaded = new Dictionary<string, byte[]>()
					{
						{ fileName, input.ToArray() }
					};
				}

				if (uploaded.IsNullOrEmpty())
					throw new BadRequestException("No files provided");

				if (uploaded.Count > 10)
					throw new BadRequestException("Maximum 10 files allowed");

				return await process(uploaded);
			}
			catch (BadRequestException ex)
			{
				return badRequestHandler(ex, uploaded);
			}
			catch (Exception ex)
			{
				Logger.LogEmailError(ex, appName);
				return exceptionHandler(ex, uploaded);
			}
		}

		protected async Task<T> UploadAndHandle<T>(string appName, 
			Func<IDictionary<string, byte[]>, T> process, 
			Func<BadRequestException, IDictionary<string, byte[]>, T> badRequestHandler, 
			Func<Exception, IDictionary<string, byte[]>, T> exceptionHandler)
		{
			IDictionary<string, byte[]> uploaded = null;

			try
			{
				uploaded = await UploadFiles();

				if (uploaded.IsNullOrEmpty())
					throw new BadRequestException("No files provided");

				if (uploaded.Count > 10)
					throw new BadRequestException("Maximum 10 files allowed");

				return process(uploaded);
			}
			catch (BadRequestException ex)
			{
				return badRequestHandler(ex, uploaded);
			}
			catch (Exception ex)
			{
				Logger.LogEmailError(ex, appName);
				return exceptionHandler(ex, uploaded);
			}
		}

		protected async Task<T> UploadAndHandle<T>(string appName, 
			Func<IDictionary<string, byte[]>, Task<T>> process, 
			Func<BadRequestException, IDictionary<string, byte[]>, T> badRequestHandler, 
			Func<Exception, IDictionary<string, byte[]>, T> exceptionHandler)
		{
			IDictionary<string, byte[]> uploaded = null;

			try
			{
				uploaded = await UploadFiles();
				return await process(uploaded);
			}
			catch (BadRequestException ex)
			{
				return badRequestHandler(ex, uploaded);
			}
			/*catch (FormatNotSupportedException ex)
			{
				return badRequestHandler(ex, uploaded.folderName, uploaded.fileNames);
			}
			catch (NotSupportedException ex)
			{
				return badRequestHandler(ex, uploaded.folderName, uploaded.fileNames);
			}*/
			catch (Exception ex)
			{
				Logger.LogEmailError(ex, appName);
				return exceptionHandler(ex, uploaded);
			}
		}

		protected byte[] ReadAndRemoveAsBytes(IDictionary<string, byte[]> files, string ending)
		{
			var keyFile = files.FirstOrDefault(x => x.Key.EndsWith(ending));

			if (keyFile.Value != null)
			{

				files.Remove(keyFile.Key);
				return keyFile.Value;
			}

			return null;
		}

		protected string ReadAndRemoveAsText(IDictionary<string, byte[]> files, string ending)
		{
			var keyFile = files.FirstOrDefault(x => x.Key.EndsWith(ending));

			if (keyFile.Value != null)
			{
				files.Remove(keyFile.Key);
				return Encoding.UTF8.GetString(keyFile.Value);
			}

			return null;
		}
	}
}
