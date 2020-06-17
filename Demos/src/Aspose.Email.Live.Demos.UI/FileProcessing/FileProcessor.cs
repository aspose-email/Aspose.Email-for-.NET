using  Aspose.Email.Live.Demos.UI.Models;
using System.Diagnostics;
using System;

namespace Aspose.Email.Live.Demos.UI.FileProcessing
{
	///<Summary>
	/// FileProcessor class to process file
	///</Summary>
	public abstract class FileProcessor
    {
		///<Summary>
		/// FileProcessor class process method
		///</Summary>
		///<param name="inputFolderName"></param>
		///<param name="inputFileName"></param>
		///<param name="outputFileName"></param>
		public virtual Response Process(string inputFolderName, string inputFileName, string outputFileName = null)
        {
            var resp = new Response()
            {
                StatusCode = 200,
                Status = "OK"
            };

            ProcessFileToResponce(resp, inputFolderName, inputFileName, outputFileName);

            return resp;
        }
		
		///<Summary>
		/// Process files in input and modify Response 
		///</Summary>
		///<param name="inputFolderName"></param>
		///<param name="inputFileNames">Short file names</param>
		///<param name="resp">Responce to modify</param>
		protected abstract void ProcessFileToResponce(Response resp, string inputFolderName, string[] inputFileNames);
		///<Summary>
		/// Process
		///</Summary>
		/// <param name="inputFolderName"></param>
		/// <param name="inputFileName"></param>
		/// <param name="outFileName"></param>
		public   Response Process(string inputFolderName, string[] inputFileNames)
		{
			var stackTrace = new StackTrace();

			string methodName = null;
			string ModelName = null;

			for (int i = 1; i < stackTrace.FrameCount; i++)
			{
				var method = stackTrace.GetFrame(i).GetMethod();

				if (method.DeclaringType != null && typeof(EmailBase).IsAssignableFrom(method.DeclaringType))
				{
					methodName = method.Name;
					ModelName = method.DeclaringType.Name;
					break;
				}
				else
				{
					methodName = method.Name;
					ModelName = "NULL";
				}
			}
			try
			{
				var resp = new Response()
				{
					StatusCode = 200,
					Status = "OK"
				};

				ProcessFileToResponce(resp, inputFolderName, inputFileNames);

				return resp;
			}
			catch (Exception ex)
			{
				var resp = new Response()
				{
					StatusCode = 500,
					Status = ex.Message
				};

				
				return resp;
			}
		}
		///<Summary>
		/// FileProcessor class ProcessFileToResponce method
		///</Summary>
		///<param name="inputFolderName"></param>
		///<param name="inputFileName"></param>
		///<param name="outputFileName"></param>
		///<param name="resp"></param>
		protected abstract void ProcessFileToResponce(Response resp, string inputFolderName, string inputFileName, string outputFileName = null);
    }
}
