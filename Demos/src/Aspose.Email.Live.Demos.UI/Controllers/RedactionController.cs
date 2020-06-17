using Aspose.Email.Live.Demos.UI.Models.Common;
using Aspose.Email.Live.Demos.UI.Models;
using Aspose.Email.Live.Demos.UI.Services;
using System;
using System.Collections;
using System.Web;
using System.Web.Mvc;

namespace Aspose.Email.Live.Demos.UI.Controllers
{
	public class RedactionController : BaseController  
	{
		public override string Product => (string)RouteData.Values["product"];


		[HttpPost]
		public Response Redaction(string outputType, string searchQuery, string replaceText, bool caseSensitive, bool text, bool comments, bool metadata)
		{
			Response response = null;
			if (Request.Files.Count > 0)
			{
				string _sourceFolder = Guid.NewGuid().ToString();
				var docs = UploadFiles(Request, _sourceFolder);

				AsposeEmailRedaction asposeEmailRedaction = new AsposeEmailRedaction();
				response = asposeEmailRedaction.Redact(docs, outputType, searchQuery, replaceText, caseSensitive, text, comments, metadata);

			}

			return response;				
		}
		public ActionResult Redaction()
		{
			var model = new ViewModel(this, "Redaction")
			{
				ControlsView = "RedactionControls",
				
				MaximumUploadFiles = 10,
				DropOrUploadFileLabel = Resources["DropOrUploadFiles"]
			};
			if (model.RedirectToMainApp)
				return Redirect("/email/" + model.AppName.ToLower());
			return View(model);			
		}	

	}
}
