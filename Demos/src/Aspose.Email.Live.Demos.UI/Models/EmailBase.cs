using System;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Aspose.Email.Live.Demos.UI.Models;
using Aspose.Email.Live.Demos.UI.Services;
using File = System.IO.File;
using Aspose.Email.Live.Demos.UI.Helpers;
using Aspose.Email.Live.Demos.UI.FileProcessing;

namespace Aspose.Email.Live.Demos.UI.Models
{
	///<Summary>
	/// EmailBase class 
	///</Summary>

	public class EmailBase : ModelBase
	{
		
		/// <summary>
		/// Maximum number of files which can be uploaded for MVC Aspose.Words apps
		/// </summary>
		protected const int MaximumUploadFiles = 10;

    /// <summary>
    /// Original file format SaveAs option for multiple files uploading. By default, "-"
    /// </summary>
    protected const string SaveAsOriginalName = ".-";
    
    /// <summary>
    /// Response when uploaded files exceed the limits
    /// </summary>
    protected Response MaximumFileLimitsResponse = new Response()
    {
      Status = $"Number of files should be less {MaximumUploadFiles}",
      StatusCode = 500
    };
		/// <summary>
		/// Response when uploaded files exceed the limits
		/// </summary>
		protected Response PasswordProtectedResponse = new Response()
		{
			Status = "Some of your documents are password protected",
			StatusCode = 500
		};
		/// <summary>
		/// Response when uploaded files exceed the limits
		/// </summary>
		protected Response BadDocumentResponse = new Response()
		{
			Status = "Some of your documents are corrupted",
			StatusCode = 500
		};
		

		///<Summary>
		/// Aspose Email Options Class
		///</Summary>
		protected class Options
		{
			///<Summary>
			/// AppName
			///</Summary>
			public string AppName;

			///<Summary>
			/// FolderName
			///</Summary>
			public string FolderName;

			///<Summary>
			/// FileName
			///</Summary>
			public string FileName;

			private string _outputType;

			/// <summary>
			/// By default, it is the extension of FileName
			/// </summary>
			public string OutputType
			{
				get => _outputType;
				set
				{
					if (!value.StartsWith("."))
						value = "." + value;
					_outputType = value;
				}
			}

			/// <summary>
			/// Check if OuputType is a picture extension
			/// </summary>
			public bool IsPicture
			{
				get
				{
					switch (_outputType.ToLower())
					{
						case ".bmp":
						case ".png":
						case ".jpg":
						case ".jpeg":
							return true;
						default:
							return false;
					}
				}
			}

			///<Summary>
			/// ResultFileName
			///</Summary>
			public string ResultFileName;

			///<Summary>
			/// MethodName
			///</Summary>
			public string MethodName;

			///<Summary>
			/// ModelName
			///</Summary>
			public string ModelName;

			///<Summary>
			/// CreateZip
			///</Summary>
			public bool CreateZip;

			///<Summary>
			/// CheckNumberOfPages
			///</Summary>
			public bool CheckNumberOfPages = false;

			///<Summary>
			/// DeleteSourceFolder
			///</Summary>
			public bool DeleteSourceFolder = false;

			///<Summary>
			/// CalculateZipFileName
			///</Summary>
			public bool CalculateZipFileName = true;

			/// <summary>
			/// Output zip filename (without '.zip'), if CreateZip property is true
			/// By default, FileName + AppName
			/// </summary>
			public string ZipFileName;

			
			/// <summary>
			/// AppSettings.WorkingDirectory + FolderName + "/" + FileName
			/// </summary>
			public string WorkingFileName
			{
				get
				{
					if (File.Exists(Config.Configuration.WorkingDirectory + FolderName + "/" + FileName))
						return Config.Configuration.WorkingDirectory + FolderName + "/" + FileName;
					return Config.Configuration.OutputDirectory + FolderName + "/" + FileName;
				}
			}

			

		}
		protected string ReadAndRemoveAsText(ref string[] files, string folder, string key)
		{
			if (files.Contains(key))
			{
				files = files.Except(new[] { key }).ToArray();
				var path = Path.Combine(Config.Configuration.WorkingDirectory, folder, key);
				var data = File.ReadAllText(path);
				return data;
			}

			return null;
		}

		protected byte[] ReadAndRemoveAsBytes(ref string[] files, string folder, string key)
		{
			if (files.Contains(key))
			{
				files = files.Except(new[] { key }).ToArray();
				//var path = Path.Combine(Config.Configuration.WorkingDirectory, folder, key);
				var data = File.ReadAllBytes(key);
				return data;
			}

			return null;
		}

		protected string GetFilePathWithName(string[] files, string folder, string key)
		{
			if (files.Contains(key))
			{
				var path = Path.Combine(Config.Configuration.WorkingDirectory, folder, key);
				return path;
			}

			return null;
		}
		/// <summary>
		/// Prepare upload files and return as folder and names
		/// </summary>
		protected async Task<(string, string[])> UploadFiles()
		{
			var folderName = Guid.NewGuid().ToString();
			var pathProcessor = new PathProcessor(folderName);
			var uploadProvider = new MultipartFormDataStreamProviderSafe(pathProcessor.SourceFolder);

			await Request.Content.ReadAsMultipartAsync(uploadProvider);

			string[] rawContentFiles = new string[0];

			if (uploadProvider.Contents?.Count > 0)
			{
				rawContentFiles = await Task.WhenAll(uploadProvider.Contents.Where(x =>
				{
					return x.Headers?.ContentDisposition?.FileName == null;
				}).Select(async x =>
				{
					try
					{
						using (var stream = await x.ReadAsStreamAsync())
						{
							var name = x.Headers?.ContentDisposition?.Name?.Trim('"') ?? Guid.NewGuid().ToString();
							var path = Path.Combine(pathProcessor.SourceFolder, name);
							using (var file = System.IO.File.Create(path))
								await stream.CopyToAsync(file);
							return name;
						}
					}
					catch (Exception ex)
					{
						Console.WriteLine(ex.Message);
						throw;
					}
				}).ToArray());
			}

			return (folderName, uploadProvider.FileData.Select(x => x.LocalFileName).Concat(rawContentFiles).ToArray());
		}
		/// <summary>
		/// init Options
		/// </summary>
		protected Options Opts = new Options();
    
    /// <summary>
    /// UTF8WithoutBom
    /// </summary>
    protected static readonly Encoding UTF8WithoutBom = new UTF8Encoding(false);

		

		

		



		/// <summary>
		/// Set default parameters into Opts
		/// </summary>
		/// <param name="filename"></param>
		private void SetDefaultOptions(string filename, string outputType)
		{
			//Opts.FolderName = FolderName;
			Opts.ResultFileName = filename;
			Opts.FileName = Path.GetFileName(filename);

			//var query = Request.GetQueryNameValuePairs().ToDictionary(kv => kv.Key, kv => kv.Value, StringComparer.OrdinalIgnoreCase);

			//if (query.ContainsKey("outputType"))
			//outputType = query["outputType"];
			Opts.OutputType = !string.IsNullOrEmpty(outputType)
			  ? outputType
			  : Path.GetExtension(Opts.FileName);

			Opts.ResultFileName = Opts.OutputType == SaveAsOriginalName
			  ? Opts.FileName
			  : Path.GetFileNameWithoutExtension(Opts.FileName) + Opts.OutputType;
		}

		

		/// <summary>
		/// Process
		/// </summary>
		protected Response Process(ActionDelegate action)
		{
			if (string.IsNullOrEmpty(Opts.OutputType))
				Opts.OutputType = Path.GetExtension(Opts.FileName);
			if (Opts.OutputType.ToLower() == ".html" || Opts.OutputType == ".SVG" || Opts.IsPicture)
				Opts.CreateZip = true;
			if (string.IsNullOrEmpty(Opts.ZipFileName) && Opts.CalculateZipFileName)
				Opts.ZipFileName = Path.GetFileNameWithoutExtension(Opts.FileName) + Opts.AppName;


			return Process(GetType().Name, Opts.ResultFileName, Opts.FolderName, Opts.OutputType, Opts.CreateZip,
				Opts.CheckNumberOfPages,
				 Opts.MethodName, action,
				Opts.DeleteSourceFolder, Opts.ZipFileName);
		}

		

		

		

		

		
	}
}
