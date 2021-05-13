using Aspose.Email.Live.Demos.UI.Models;
using Aspose.Email.Live.Demos.UI.Config;
using Newtonsoft.Json;
using System.Collections.Generic;

namespace Aspose.Email.Live.Demos.UI.Models
{
	public class JSOptions
	{
		public string AppURL { get; }
		public string AppName { get; }
		public string AppRoute { get; }

		public string APIBasePath { get; }
		public string UIBasePath { get; }

		public string ViewerPathWF { get; }

		public string ViewerPath { get; }
		public string EditorPath { get; }

		public string FileSelectMessage { get; }

		public string FirstNameInputMessage { get; }
		public string PasswordInputMessage { get; }

		public Dictionary<string, string> Keys { get; }

		public bool ShowViewerButton { get; }
		public int MaximumUploadFiles { get; }

		public string FileAmountMessage { get; }

		/// <summary>
		/// Apps like Viewer and Editor
		/// </summary>
		public bool UploadAndRedirect { get; }
		public bool NeedsProcessButton { get; }
		public bool UseSorting { get; }

		public bool ShowButtonBar { get; }

		public string FileWrongTypeMessage { get; }

		public Dictionary<int, string> FileProcessingErrorCodes { get; }

		public IEnumerable<string> UploadOptions { get; }

		#region FileDrop
		public bool Multiple { get; }
		public string DropFilesPrompt { get; }
		public string Accept { get; }
		public bool CanWorkWithoutFiles { get; }
		public bool DefaultFileBlockDisabled { get; }
		#endregion

		public JSOptions(ViewModel model)
		{
			AppURL = model.AppURL;
			AppName = model.AppName;
			AppRoute = model.AppRoute;

			APIBasePath = $"{Configuration.WebApiPath}api/";

			UIBasePath = model.BaseUrl;

			ShowButtonBar = model.ShowFileDropButtonBar;
			UseSorting = model.UseSorting;
			UploadAndRedirect = model.UploadAndRedirect;
			NeedsProcessButton = model.NeedsProcessButton;
			FileAmountMessage = model["FileAmountMessageLessTen"];
			MaximumUploadFiles = ViewModel.MaximumUploadFiles;
			ShowViewerButton = model.ShowViewerButton;
			FirstNameInputMessage = model["FirstNameInputMessage"];
			PasswordInputMessage = model["PasswordInputMessage"];

			Keys = new Dictionary<string, string>
			{
				{ "NumberOfCharacters", model["NumberOfCharacters"] },
				{ "UppercaseLetters", model["UppercaseLetters"] },
				{ "LowercaseLetters", model["LowercaseLetters"] },
				{ "Numbers", model["Numbers"] },
				{ "Symbols", model["Symbols"] },
				{ "MiddleNumbersOrSymbols", model["MiddleNumbersOrSymbols"] },
				{ "Requirements", model["Requirements"] },
				{ "LettersOnly", model["LettersOnly"] },
				{ "NumbersOnly", model["NumbersOnly"] },
				{ "RepeatCharacters", model["RepeatCharacters"] },
				{ "ConsecutiveUppercaseLetters", model["ConsecutiveUppercaseLetters"] },
				{ "ConsecutiveLowercaseLetters", model["ConsecutiveLowercaseLetters"] },
				{ "ConsecutiveNumbers", model["ConsecutiveNumbers"] },
				{ "SequentialLetters3Plus", model["SequentialLetters3Plus"] },
				{ "SequentialNumbers3Plus", model["SequentialNumbers3Plus"] },
				{ "SequentialSymbols3Plus", model["SequentialSymbols3Plus"] },
			};

			FileSelectMessage = model["FileSelectMessage"];
			EditorPath = $"{UIBasePath}/{model.Product}/edit?";
			ViewerPath = $"{UIBasePath}/{model.Product}/view?";
			ViewerPathWF = $"{UIBasePath}/{model.Product}/viewer/";

			model.ExtensionsString.Replace(".", "").ToUpper().Split('|');
			UploadOptions = model.ExtensionsString.Replace(".", "").ToUpper().Split('|');
			DropFilesPrompt = model["DropOrUploadFiles"];
			Accept = model.ExtensionsString.Replace("|.", ",.");
			CanWorkWithoutFiles = model.CanWorkWithoutFiles;
			DefaultFileBlockDisabled = model.DefaultFileBlockDisabled;

			Multiple = !UploadAndRedirect;

			FileProcessingErrorCodes = new Dictionary<int, string>()
			{
				{ (int)FileProcessingErrorCode.NoSearchResults, model["NoSearchResultsMessage"] },
				{ (int)FileProcessingErrorCode.WrongRegExp, model["WrongRegExpMessage"] }
			};

            if (string.IsNullOrEmpty(model.Extension1) || model.IsCanonical)
                FileWrongTypeMessage = model["FileWrongTypeMessage"];
            else
                FileWrongTypeMessage = string.Format(model["FileWrongTypeMessage2"], $"<a href=\"/{model.Product}/{AppName.ToLower()}\">{AppName}</a>");
		}

		public override string ToString()
		{
			return JsonConvert.SerializeObject(this, Formatting.None);
		}
	}
}
