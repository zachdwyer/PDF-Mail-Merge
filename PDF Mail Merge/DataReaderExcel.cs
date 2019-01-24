using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace PDF_Mail_Merge
{
	class DataReaderExcel : DataReader
	{
		/// <summary>
		/// Stores meta data regarding each column in the spreadsheet
		/// </summary>
		private Dictionary<String, int> columnMeta;

		/// <summary>
		/// Default Constructor
		/// </summary>
		public DataReaderExcel()
			: base()
		{
		}

		/// <summary>
		/// Construct the object and specify a file path for a file to be parsed.
		/// </summary>
		/// <param name="filePath">Path of the file to be parsed.</param>
		public DataReaderExcel(String filePath)
			: base(filePath)
		{
		}

		/// <summary>
		/// Construct the object and specify the dataset to work with.
		/// </summary>
		/// <param name="dataSet">The dataset.</param>
		public DataReaderExcel(DataTable dataSet)
			: base(dataSet)
		{
		}

		/// <summary>
		/// Contruct the object and specify both a file path of a file to be parsed and a dataset to work with.
		/// </summary>
		/// <param name="filePath">Path of the file to be parsed.</param>
		/// <param name="dataSet">The dataset.</param>
		public DataReaderExcel(String filePath, DataTable dataSet)
			: base(filePath, dataSet)
		{
		}

		/// <summary>
		/// Initiate a parse of the loaded file.
		/// </summary>
		/// <returns>Success of the parse operation.</returns>
		public override bool parseFile()
		{
			return parseFile("");
		}

		/// <summary>
		/// Initiate a parse of the loaded file.
		/// </summary>
		/// <param name="sheetName">The name of the Sheet to parse from the Excel Workbook</param>
		/// <returns>Success of the parse operation.</returns>
		public bool parseFile(String sheetName)
		{
			// Open Excel Workbook (as a DB)
			//OleDbConnection workbook = new OleDbConnection("provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + getFilePath() + "';Extended Properties='Excel 12.0 Xml;HDR=YES;ReadOnly=TRUE;IMEX=1'");
			OleDbConnection workbook = new OleDbConnection("provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + getFilePath() + "';Extended Properties='Excel 12.0 Xml;HDR=YES;ReadOnly=TRUE'");
			workbook.Open();

			// Make a temporary datatable
			DataTable dt = workbook.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

			// Create an array of sheet names
			String[] excelSheets = new String[dt.Rows.Count];
			int i = 0;
			foreach (DataRow row in dt.Rows)
			{
				excelSheets[i] = row["TABLE_NAME"].ToString();

				// Un-escape sheet names
				excelSheets[i] = excelSheets[i].Replace("''", "'");
				if (excelSheets[i].IndexOf(' ') > -1)
				{
					excelSheets[i] = excelSheets[i].Substring(1, excelSheets[i].Length - 2);
				}
				// Remove trailing $
				excelSheets[i] = excelSheets[i].Substring(0, excelSheets[i].Length - 1);

				i++;
			}

			if (sheetName == null || sheetName.Length == 0)
			{
				// Prompt user for which sheet to query against
				DataReaderExcelSheetPrompt prompt = new DataReaderExcelSheetPrompt(excelSheets);

				if (prompt.ShowDialog() == DialogResult.OK)
				{
					sheetName = prompt.SelectedSheet;
				}
			}
			
			if( !excelSheets.Contains(sheetName) )
			{
				// Throw an error
			}
			else
			{
				// Add trailing $ to the name of the sheet
				sheetName += "$";

				// Escape sheet name
				//sheet = sheet.Replace("'", "''");

				if (sheetName.IndexOf(' ') > -1)
				{
					// Add single quotes back if necessary
					sheetName = "'" + sheetName + "'";
				}

				// Get column meta data
				try
				{
					dt = workbook.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, null);
				}
				catch (System.Data.OleDb.OleDbException e)
				{
					// Return error code
					return false;
				}

				this.columnMeta = new Dictionary<String, int>();
				int j = 0;
				foreach (DataRow row in dt.Rows)
				{
					if (row["TABLE_NAME"].ToString() == sheetName)
					{
						// Store the data type into the column meta data variable
						this.columnMeta[row["COLUMN_NAME"].ToString()] = (int)row["DATA_TYPE"];

						// Reset the column data type to "string", for formatting purposes
						//dt.Rows[i]["DATA_TYPE"] = OleDbType.VarChar;

					}

					++j;
				}

				// Query the sheet
				OleDbDataAdapter command = new OleDbDataAdapter("select * from [" + sheetName + "]", workbook);

				// Name the table
				//command.TableMappings.Add("Table", "MailMerge");

				// Store the data from the workbook into a dataSet
				dt = new DataTable();

				command.Fill(dt);

				this.fillableDataSet = dt.Clone();
				for (int column = 0; column < this.fillableDataSet.Columns.Count; column++)
				{
					this.fillableDataSet.Columns[column].DataType = typeof(String);
				}

				foreach (DataRow row in dt.Rows)
				{
					//DataRow dr = new DataRow();
					DataRow dr = this.fillableDataSet.NewRow();
					foreach (String field in columnMeta.Keys)
					{
						dr[field] = this.formatData(field, row[field]);
					}
					this.fillableDataSet.Rows.Add(dr);

					//this.fillableDataSet.ImportRow(row);
				}

				//dt = workbook.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

				// Close the workbook
				workbook.Close();

				// Free up memory & The File
				workbook.Dispose();

				return true;
			}

			return false;
		}

		/// <summary>
		/// Convert an OLEDB data object from a sheet in the workbook to a string. 
		/// Format the string conditionally on the type of the data object.
		/// </summary>
		/// <param name="field">Name of the column in the spreadsheet</param>
		/// <param name="data">Object data from OLEDB data object</param>
		/// <returns></returns>
		protected String formatData(String field, Object data)
		{
			String formattedString = "";

			if (data.ToString().Length > 0 && data.ToString().ToUpper() != "NULL")
			{
				formattedString = data.ToString();

				// Format the output based on the column type
				switch ((OleDbType)columnMeta[field])
				{
					// Strings - Don't need to process because of ToString() above
					case OleDbType.BSTR:
					// DBTYPE_BSTR
					case OleDbType.Char:
					// DBTYPE_STR
					case OleDbType.WChar:
					// DBTYPE_WSTR
					case OleDbType.VarChar:
					case OleDbType.LongVarChar:
					case OleDbType.VarWChar:
					case OleDbType.LongVarWChar:
						break;

					// Numbers - Don't need to process because of ToString() above
					case OleDbType.SmallInt:
					// DBTYPE_I2
					case OleDbType.Integer:
					// DBTYPE_I4
					case OleDbType.Single:
					// DBTYPE_R4
					case OleDbType.Double:
					// DBTYPE_R8
					case OleDbType.Decimal:
					// DBTYPE_DECIMAL
					case OleDbType.TinyInt:
					// DBTYPE_I1
					case OleDbType.UnsignedTinyInt:
					// DBTYPE_UI1
					case OleDbType.UnsignedSmallInt:
					// DBTYPE_UI2
					case OleDbType.UnsignedInt:
					// DBTYPE_UI4
					case OleDbType.BigInt:
					// DBTYPE_UI4
					case OleDbType.UnsignedBigInt:
					// DBTYPE_UI8
					case OleDbType.Numeric:
					// DBTYPE_NUMERIC
					case OleDbType.VarNumeric:
						break;

					// Dates
					case OleDbType.Date:
						// DBTYPE_DATE - Date format mm/dd/yyyy
						formattedString = ((DateTime)data).ToShortDateString();
						break;
					case OleDbType.DBDate:
						// DBTYPE_DBDATE - Date format yyyymmdd
						formattedString = ((DateTime)data).ToShortDateString();
						break;
					case OleDbType.DBTime:
						// DBTYPE_DBTIME - Date format hhmmss
						formattedString = ((DateTime)data).ToShortTimeString();
						break;
					case OleDbType.DBTimeStamp:
						// DBTYPE_DBTIMESTAMP - Date format yyyymmddhhmmss
						formattedString = ((DateTime)data).ToString();
						break;
					case OleDbType.Filetime:
						// DBTYPE_FILETIME - Date format to 100 nanosecond intervals since Jan 1, 1601
						break;

					// Currency
					case OleDbType.Currency:
						// DBTYPE_CY - Money Format
						formattedString = string.Format("{0:C}", data);
						break;
					/*
					// None of these types should be available to Excel
					case OleDbType.Boolean:
						// DBTYPE_BOOL
						break;

					case OleDbType.Variant:
						// DBTYPE_VARIANT
						break;

					case OleDbType.Guid:
						// DBTYPE_GUID
						break;

					case OleDbType.Binary:
						// DBTYPE_BYTES
					case OleDbType.VarBinary:
					case OleDbType.LongVarBinary:
						break;

					case OleDbType.Empty:
						// DBTYPE_EMPTY
						break;

					case OleDbType.Error:
						// DBTYPE_ERROR
					case OleDbType.IUnknown:
						// DBTYPE_UNKNOWN
					case OleDbType.IDispatch:
						// DBTYPE_IDISPATCH
					*/
					default:
						//MessageBox.Show( data.GetType().ToString() );
						break;
				}
			}

			return formattedString;
		}

		/// <summary>
		/// Get the file type filter for a browse for file button.
		/// </summary>
		public override String getFileFilter()
		{
			return "Excel Workbook (*.xlsx)|*.xlsx";
		}

		/// <summary>
		/// Get the file extension list that this DataReader parses.
		/// </summary>
		public override String[] getFileFilterExtensions()
		{
			return new String[] { ".xlsx", ".xls" };
		}
	}
}
