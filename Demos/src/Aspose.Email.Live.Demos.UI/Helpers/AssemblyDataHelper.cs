using Aspose.Cells;
using Microsoft.VisualBasic.FileIO;
using System;
using System.Data;
using System.IO;
using System.Xml;

namespace Aspose.Email.Live.Demos.UI.Helpers
{
    public static class AssemblyDataHelper
	{
		public static DataTable PrepareDataTable(string filename, Stream input, string datasourceName, int datasourceTableIndex = 0, string delimiter = ",")
		{
			DataTable table;

			var format = Path.GetExtension(filename).ToLower();

			switch (format)
			{
				case ".json":
					table = PrepareDataTableFromJson(input, datasourceName);
					break;
				case ".xml":
					table = PrepareDataTableFromXML(input, datasourceName);
					break;
				case ".csv":
					table = PrepareDataTableFromCSV(input, datasourceName, delimiter);
					break;
				case ".xls":
				case ".xlsx":
					table = PrepareDataTableFromExcel(input, datasourceTableIndex);
					break;
				default:
					throw new NotSupportedException("Invalid data file format: " + format);
			}

			return table;
		}

		public static DataTable PrepareDataTableFromJson(Stream input, string datasourceName)
		{
			DataSet dataSet = input.ReadAs<DataSet>();
			return dataSet.Tables[datasourceName];
		}

		public static DataTable PrepareDataTableFromXML(Stream input, string datasourceName)
		{
			var dataSet = new DataSet("Temp");
			dataSet.ReadXml(XmlReader.Create(input));
			return dataSet.Tables[datasourceName];
		}

		public static DataTable PrepareDataTableFromCSV(Stream input, string datasourceName, string delimiter)
		{
			var dataTable = new DataTable(datasourceName);
			using (var csvReader = new TextFieldParser(input))
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

		public static DataTable PrepareDataTableFromExcel(Stream input, int datasourceTableIndex)
		{
			var excel = new Workbook(input);
			var cells = excel.Worksheets[datasourceTableIndex].Cells;
			var lastColumn = cells.MaxColumn;
			var lastRow = int.MinValue;
			for (int i = 0; i < lastColumn; i++)
				if (cells.GetLastDataRow(i) > lastRow)
					lastRow = cells.GetLastDataRow(i);
			if (lastRow == int.MinValue)
				return null;
			return cells.ExportDataTable(0, 0, lastRow + 1, lastColumn + 1, true);
		}
	}
}
