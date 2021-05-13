using Aspose.Email.Live.Demos.UI.Config;
using Microsoft.Extensions.Primitives;
using System.Text;
using System.Web;

namespace Aspose.Email.Live.Demos.UI.Pages
{
    public class EditorModel : BaseRootPageModel
    {
        public override string AppName => "Editor";

        public string FileName { get; private set; }
        public string FolderName { get; private set; }
        public string CallbackURL { get; private set; }
        public string DownloadOriginalURL { get; private set; }
        public string ProductName { get; private set; }
        public string AsposeEditorApi { get; private set; }
        public string FileDownloadLink { get; private set; }
        public string DownloadDropDownItemsHtml { get; private set; }

        public override void OnGet()
        {
            base.OnGet();

            AsposeEditorApi = Configuration.WebApiPath + "api/AsposeEmailEditor/";
            FileDownloadLink = Configuration.WebApiFileDownloadPath;

            if (!this.IsPostBack())
            {
                ProductName = this["APPName"];
                FileName = Request.Query["FileName"];
                FolderName = Request.Query["FolderName"];

                if (Request.Query["callbackURL"] != StringValues.Empty)
                    CallbackURL = Request.Query["callbackURL"];
                else
                    CallbackURL = "email/editor";

                var url = new StringBuilder(Configuration.WebApiFileDownloadPath);
                url.Append("?FileName=");
                url.Append(HttpUtility.UrlEncode(FileName));
                url.Append("&FolderName=");
                url.Append(FolderName);
                DownloadOriginalURL = url.ToString();
            }

            var downloadItemsBuilder = new StringBuilder();

            downloadItemsBuilder.AppendLine(CreateListItem("MSG"));
            downloadItemsBuilder.AppendLine(CreateListItem("EML"));
            downloadItemsBuilder.AppendLine(CreateListItem("HTML"));

            DownloadDropDownItemsHtml = downloadItemsBuilder.ToString();
        }

        string CreateListItem(string extension)
        {
            return $@"<li><a id=""download-as-{extension.ToLowerInvariant()}"" href=""#"" ng-click=""Download('.{extension.ToLowerInvariant()}')"">as {extension.ToUpperInvariant()}</a></li>";
        }
    }
}
