
using Microsoft.VisualBasic.FileIO;
using Newtonsoft.Json;
using System;
using System.Data;
using System.IO;
using System.Xml;

namespace Aspose.Email.Live.Demos.UI.LibraryHelpers
{
	public static class AssemblyDataHelper
	{
		

		private static DataTable PrepareDataTableFromDocument(string filename, string datasourceName, int datasourceTableIndex)
		{
			//var data = new Document(filename);

			////var table = (Table)data.GetChild(NodeType.Table, datasourceTableIndex, true);
			////if (table == null)
			////    return null;
			////var properties = table.FirstRow.Cells.Select(x => x.GetText().Replace("\a", "")).ToArray();

			//var dataTable = new DataTable(datasourceName);
			////foreach (var property in properties)
			////    if (!dataTable.Columns.Contains(property))
			////        dataTable.Columns.Add(property);

			////foreach (var row in table.Rows.Skip(1).Select(x => (Row)x))
			////    dataTable.Rows.Add(row.Cells.Select(x => x.GetText().Replace("\a", "") as object).ToArray()); // Cells have special symbol '\a'

			//return dataTable;
			throw new NotImplementedException();
		}

		public static DataTable PrepareDataTableFromJson(string filename, string datasourceName)
		{
			string content = System.IO.File.ReadAllText(filename);
			DataSet dataSet = JsonConvert.DeserializeObject<DataSet>(content);
			return dataSet.Tables[datasourceName];
		}

		public static DataTable PrepareDataTableFromXML(string filename, string datasourceName)
		{
			var dataSet = new DataSet("Temp");
			dataSet.ReadXml(XmlReader.Create(filename));
			return dataSet.Tables[datasourceName];
		}

		public static DataTable PrepareDataTableFromCSV(string filename, string datasourceName, string delimiter)
		{
			var dataTable = new DataTable(datasourceName);
			using (var csvReader = new TextFieldParser(filename))
			{
				csvReader.SetDelimiters(delimiter);
				csvReader.HasFieldsEnclosedInQuotes = true;
				var colFields = csvReader.ReadFields();
				foreach (var column in colFields)
				{
					var datecolumn = new DataColumn(column)
					{
						AllowDBNull = true
					};
					dataTable.Columns.Add(datecolumn);
				}
				while (!csvReader.EndOfData)
				{
					var fieldData = csvReader.ReadFields();
					//Making empty value as null
					for (int i = 0; i < fieldData.Length; i++)
						if (fieldData[i] == "")
							fieldData[i] = null;
					dataTable.Rows.Add(fieldData);
				}
			}
			return dataTable;
		}

		
	}
}
