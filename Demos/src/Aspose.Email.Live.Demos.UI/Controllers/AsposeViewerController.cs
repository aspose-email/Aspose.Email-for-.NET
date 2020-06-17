using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.IO.Compression;
using System.Net;
using System.Net.Http;
using Aspose.Words;
using Aspose.Email;
using System.Net.Http.Headers;
using System.Web.Http;
using Aspose.Email.Live.Demos.UI.Helpers;
using Aspose.Email.Live.Demos.UI.Models;

namespace Aspose.Email.Live.Demos.UI.Controllers
{
	///<Summary>
	/// AsposeViewerController class to get document page
	///</Summary>
	public class AsposeViewerController : ApiController
	{
		
		///<Summary>
		/// GetDocumentPage method to get document page
		///</Summary>
		[HttpGet]		
		public HttpResponseMessage GetDocumentPage( string file, string folderName,  int currentPage)
		{
			string outfileName = Path.GetFileNameWithoutExtension(file) + "_{0}";
			string outPath = Config.Configuration.OutputDirectory + folderName + "/" + outfileName;

			currentPage = currentPage - 1;

			string imagePath = string.Format(outPath, currentPage) + ".jpeg";
			Directory.CreateDirectory(Config.Configuration.OutputDirectory + folderName);

			if (System.IO.File.Exists(imagePath))
			{
				return GetImageFromPath(imagePath);
			}
			int i = currentPage;
			return null;

			
		}
		
		///<Summary>
		/// DocumentPages method to get document pages
		///</Summary>
		[HttpGet]		
		public List<string> DocumentPages( string file, string folderName, int currentPage)
		{			
			List<string> output;			
			
			
			try
			{
				output = GetDocumentPages( file, folderName, currentPage);

				
			}
			catch (Exception ex)
			{
				
				throw ex;
			}

			return output;
		}

		private List<string> GetDocumentPages(string file, string folderName,  int currentPage)
		{
			List<string> lstOutput = new List<string>();
			string outfileName = "page_{0}";
			string outPath =  Config.Configuration.OutputDirectory + folderName + "/" + outfileName;

			currentPage = currentPage - 1;
			Directory.CreateDirectory(Config.Configuration.OutputDirectory + folderName);
			string imagePath = string.Format(outPath, currentPage) + ".jpeg";
						
			if (System.IO.File.Exists(imagePath) && currentPage > 0)
			{
				lstOutput.Add(imagePath);
				return lstOutput;
			}

			int i = currentPage;

			var filename = System.IO.File.Exists(Config.Configuration.WorkingDirectory + folderName + "/" + file)
				? Config.Configuration.WorkingDirectory + folderName + "/" + file
				: Config.Configuration.OutputDirectory + folderName + "/" + file;
			
			using (FilePathLock.Use(filename))
			{


				

				try
				{
					int tempPageCount = 0;
					Aspose.Email.Live.Demos.UI.Models.License.SetAsposeEmailLicense();
					Aspose.Email.Live.Demos.UI.Models.License.SetAsposeWordsLicense();

					// Load mail message
					MailMessage message = MailMessage.Load(filename);
					MemoryStream ms = new MemoryStream();
					message.Save(ms, Email.SaveOptions.DefaultMhtml);
					ms.Position = 0;

					// create an instance of LoadOptions and set the LoadFormat to Mhtml
					var loadOptions = new Aspose.Words.LoadOptions();
					loadOptions.LoadFormat = Words.LoadFormat.Mhtml;


					// create an instance of Document and load the MTHML from MemoryStream
					var document = new Aspose.Words.Document(ms, loadOptions);
					ms.Close();

					tempPageCount += document.PageCount;
					tempPageCount += message.Attachments.Count;

				 Aspose.Words.Saving.ImageSaveOptions options = new Aspose.Words.Saving.ImageSaveOptions(Aspose.Words.SaveFormat.Jpeg);
					options.PageCount = 1;

					// Save each page of the document as image.

					options.PageIndex = i;

					lstOutput.Add(tempPageCount.ToString());
					if (currentPage <= document.PageCount - 1)
					{
						document.Save(imagePath, options);
						lstOutput.Add(imagePath);
					}
					else
					{
						ms = new MemoryStream();
						int pageIndex = document.PageCount - currentPage;
						if (pageIndex < 0)
							pageIndex = pageIndex * -1;
						if (message.Attachments[pageIndex].IsEmbeddedMessage)
						{
							message.Attachments[pageIndex].Save(ms);
							ms.Position = 0;
							options.PageIndex = 0;

							document = new Aspose.Words.Document(ms, loadOptions);
							ms.Close();

							document.Save(imagePath, options);
						}
						else
						{
							message.Attachments[pageIndex].Save(imagePath);
						}

						lstOutput.Add(imagePath);
					}
				}
				catch (Exception ex)
				{
					throw ex;
				}
				return lstOutput;

			}
			
		}
		///<Summary>
		/// DownloadDocument method to download document
		///</Summary>
		[HttpGet]
		
		public HttpResponseMessage DownloadDocument(string file, string folderName, bool isImage)
		{			
			string outfileName = Path.GetFileNameWithoutExtension(file) + "_Out.zip";
			string outPath;

			if (!isImage)
			{
				if (System.IO.File.Exists(Config.Configuration.WorkingDirectory + folderName + "/" + file))
					outPath = Config.Configuration.WorkingDirectory + folderName + "/" + file;
				else
					outPath = Config.Configuration.OutputDirectory + folderName + "/" + file;
			}
			else
			{
				outPath = Config.Configuration.OutputDirectory + outfileName;
			}

			using (FilePathLock.Use(outPath))
			{
				if (isImage)
				{
					if (System.IO.File.Exists(outPath))
						System.IO.File.Delete(outPath);

					List<string> lst = GetDocumentPages(file, folderName,  1);

					if (lst.Count > 1)
					{
						int tmpPageCount = int.Parse(lst[0]);
						for (int i = 2; i <= tmpPageCount; i++)
						{
							GetDocumentPages( file, folderName,  i);
						}
					}

					ZipFile.CreateFromDirectory(Config.Configuration.OutputDirectory + folderName + "/", outPath);
				}


				if ((!System.IO.File.Exists(outPath)) || !Path.GetFullPath(outPath).StartsWith(Path.GetFullPath( System.Web.HttpContext.Current.Server.MapPath("~/Assets/"))))
				{
					var exception = new HttpResponseException(HttpStatusCode.NotFound);
					
					throw exception;
				}

				try
				{
					using (var fileStream = new FileStream(outPath, FileMode.Open, FileAccess.Read))
					{
						
						using (var ms = new MemoryStream())
						{
							fileStream.CopyTo(ms);
							var result = new HttpResponseMessage(HttpStatusCode.OK)
							{
								Content = new ByteArrayContent(ms.ToArray())
							};
							result.Content.Headers.ContentDisposition =
								new System.Net.Http.Headers.ContentDispositionHeaderValue("attachment")
								{
									FileName = (isImage ? outfileName : file)
								};
							result.Content.Headers.ContentType =
								new MediaTypeHeaderValue("application/octet-stream");

							return result;
						}
					}

				}
				catch (Exception x)
				{

					Console.WriteLine(x.Message);
				}
			}

			return null;
		}

		///<Summary>
		/// PageImage method to get page image
		///</Summary>
		[HttpGet]
		
		public HttpResponseMessage PageImage(string imagePath)
		{
			return GetImageFromPath(imagePath);
		}

		private HttpResponseMessage GetImageFromPath(string imagePath)
		{
			HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.OK);
			FileStream fileStream = new FileStream(imagePath, FileMode.Open, FileAccess.Read);
			try
			{
				System.Drawing.Image image = System.Drawing.Image.FromStream(fileStream);
				MemoryStream memoryStream = new MemoryStream();


				image.Save(memoryStream, ImageFormat.Jpeg);
				result.Content = new ByteArrayContent(memoryStream.ToArray());
				result.Content.Headers.ContentType = new MediaTypeHeaderValue("image/jpeg");
				fileStream.Close();

				return result;
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
				return null;
			}
		}
	}
}
