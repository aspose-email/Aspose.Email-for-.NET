using Aspose.Email.Live.Demos.UI.Models.Common;
using Aspose.Email.Live.Demos.UI.Models;
using Aspose.Email.Live.Demos.UI.Services;
using System;
using System.Collections;
using System.Web;
using System.Web.Mvc;

namespace Aspose.Email.Live.Demos.UI.Controllers
{
	public class ParserController : BaseController  
	{
		public override string Product => (string)RouteData.Values["product"];


		[HttpPost]
		public Response Parser(string outputType = "")
		{
			Response response = null;
			if (Request.Files.Count > 0)
			{
				string _sourceFolder = Guid.NewGuid().ToString();
				var docs = UploadFiles(Request, _sourceFolder);

				AsposeEmailParser emailParser = new AsposeEmailParser();
				response = emailParser.Parse(docs);

			}

			return response;			
				
		}

		

		public ActionResult Parser()
		{
			var model = new ViewModel(this, "Parser")
			{
				
				MaximumUploadFiles = 10,
				
				DropOrUploadFileLabel = Resources["DropOrUploadFiles"]
			};
			return View(model);
		}
		

	}
}
