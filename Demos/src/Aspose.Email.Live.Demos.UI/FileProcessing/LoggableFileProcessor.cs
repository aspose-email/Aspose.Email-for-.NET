using Aspose.Email.Live.Demos.UI.Controllers;
using Aspose.Email.Live.Demos.UI.Models;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading.Tasks;

namespace Aspose.Email.Live.Demos.UI.FileProcessing
{
    ///<Summary>
    /// LoggableFileProcessor class to log information in database
    ///</Summary>
    public abstract class LoggableFileProcessor : FileProcessor
    {
		public IConfiguration Configuration { get; }
		public ILogger Logger { get; }

		public LoggableFileProcessor(ILogger logger, IConfiguration configuration)
		{
			Logger = logger;
			Configuration = configuration;
		}

		///<Summary>
		/// Get or set product name
		///</Summary>
		public string ProductName { get; set; }

		///<Summary>
		/// Process
		///</Summary>
		/// <param name="inputFolderName"></param>
		/// <param name="inputFileName"></param>
		/// <param name="outFileName"></param>
		public override async Task<Response> Process(IDictionary<string, byte[]> files)
		{
			var stackTrace = new StackTrace();

			string methodName = null;
			string controllerName = null;

			for (int i = 1; i < stackTrace.FrameCount; i++)
			{
				var method = stackTrace.GetFrame(i).GetMethod();

				if (method.DeclaringType != null && typeof(ApiControllerBase).IsAssignableFrom(method.DeclaringType))
				{
					methodName = method.Name;
					controllerName = method.DeclaringType.Name;
					break;
				}
				else
				{
					methodName = method.Name;
					controllerName = "NULL";
				}
			}

			var logMsg = $"ControllerName = {(controllerName ?? "Unknown")}, MethodName = {(methodName ?? "Unknown")}";

			try
			{
				var resp = await base.Process(files);
				Logger.LogInformation(logMsg, ProductName);
				return resp;
			}
			
			catch (NotSupportedException ex)
			{
				return new Response()
				{
					StatusCode = 400,
					Status = ex.Message
				};
			}
			catch(FormatNotSupportedException ex)
			{
				return new Response()
				{
					StatusCode = 400,
					Status = ex.Message
				};
			}
			catch (BadRequestException ex)
			{
				return new Response()
				{
					StatusCode = 400,
					Status = ex.Message
				};
			}
			catch (Exception ex)
			{
				var resp = new Response()
				{
					StatusCode = 500,
					Status = "Error on processing: "+ex.Message
				}.WithTrace(Configuration, ex);

				Logger.LogEmailError(ex, logMsg, ProductName);
				return resp;
			}
		}
	}
}
