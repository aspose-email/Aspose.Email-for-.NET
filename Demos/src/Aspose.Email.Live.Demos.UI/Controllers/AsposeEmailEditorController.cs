using Aspose.Email.Live.Demos.UI.LibraryHelpers;
using Aspose.Email.Live.Demos.UI.Models;
using Aspose.Email;
using Aspose.Email.Mapi;
using System;
using System.IO;
using System.Web;
using File = System.IO.File;

namespace Aspose.Email.Live.Demos.UI.Controllers
{
	public class AsposeEmailEditorController : EmailBase
	{
        public string GetHTML(string fileName, string folderName)
        {
            var inputFolderPath = Path.Combine(Config.Configuration.WorkingDirectory, folderName);
            var inputFilePath = Path.Combine(inputFolderPath, fileName);

            return MapiHelper.GetMapiMessageFromFile(inputFilePath).BodyHtml;
        }

        public Response UpdateContents(string fileName, string folderName, string htmldata, string outputType)
        {
            outputType = outputType.ToLower();
            var savingFileName = Path.GetFileNameWithoutExtension(fileName) + outputType;
            var outputFolderName = Guid.NewGuid().ToString();

            var resultFile = Path.Combine(Config.Configuration.OutputDirectory, outputFolderName, savingFileName);
            Directory.CreateDirectory(Path.GetDirectoryName(resultFile));

            try
            {
                switch (outputType)
                {
                    case ".html":
                        File.WriteAllText(resultFile, htmldata);
                        break;
                    default:
                        {
                            var message = MapiHelper.GetMapiMessageFromFile(Path.Combine(Config.Configuration.WorkingDirectory, folderName, fileName));

                            message.SetBodyContent(htmldata, BodyContentType.Html);

                            if (outputType == ".eml")
                                message.Save(resultFile, EmlSaveOptions.DefaultEml);
                            else
                                message.Save(resultFile, MsgSaveOptions.DefaultMsgUnicode);
                            break;
                        }
                }

                return new Response()
                {
                    FileName = HttpUtility.UrlEncode(savingFileName),
                    FolderName = outputFolderName,
                    StatusCode = 200
                };

            }
			catch(Exception ex)
            {
                

                return new Response()
                {
                    FileName = HttpUtility.UrlEncode(savingFileName),
                    FolderName = outputFolderName,
                    StatusCode = 500,
                    Status = ex.Message
                };
            }
        }
    }
}
