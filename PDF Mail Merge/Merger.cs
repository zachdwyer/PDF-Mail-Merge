using System;
using System.Data;
using System.IO;
using iTextSharp;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.xml;

namespace PDF_Mail_Merge
{
	class Merger
	{
		/// <summary>
		/// Path to the PDF Template
		/// </summary>
		public String templatePath;

		/// <summary>
		/// Path to the output directory, where PDFs will be saved en masse
		/// </summary>
		public String destinationPath;

		/// <summary>
		/// Data to be filled into the PDF template
		/// </summary>
		public DataTable fillableData;

		/// <summary>
		/// Default Constructor
		/// </summary>
		public Merger()
			: this(null, null, null)
		{
		}

		/// <summary>
		/// Construct the merger object, specifying the path to the PDF Template
		/// </summary>
		/// <param name="templatePath">Path to the PDF Template</param>
		public Merger(String templatePath)
			: this(templatePath, null, null)
		{
		}

		/// <summary>
		/// Construct the merger object, specifying the path to the PDF Template and the output directory
		/// </summary>
		/// <param name="templatePath">Path to the PDF Template</param>
		/// <param name="destinationPath">Path to the output directory</param>
		public Merger(String templatePath, String destinationPath)
			: this(templatePath, destinationPath, null)
		{
		}

		/// <summary>
		/// Construct the merger object, specifying the path to the PDF Template and the data to be filled into the template
		/// </summary>
		/// <param name="templatePath">Path to the PDF Template</param>
		/// <param name="fillableData">Data to be filled into the PDF template</param>
		public Merger(String templatePath, DataTable fillableData)
			: this(templatePath, null, fillableData)
		{
		}

		/// <summary>
		/// Construct the merger object, specifying the data to be filled into a PDF template
		/// </summary>
		/// <param name="fillableData">Data to be filled into the PDF template</param>
		public Merger(DataTable fillableData)
			: this(null, null, fillableData)
		{
		}

		/// <summary>
		/// Construct the merger object, specifying the path to the PDF Template, the output directory, and the data to be filled into the template
		/// </summary>
		/// <param name="templatePath">Path to the PDF Template</param>
		/// <param name="destinationPath">Path to the output directory</param>
		/// <param name="fillableData">Data to be filled into the PDF template</param>
		public Merger(String templatePath, String destinationPath, DataTable fillableData)
		{
			this.setTemplatePath(templatePath);
			this.setDestinationPath(destinationPath);
			this.setFillableData(fillableData);
		}

		/// <summary>
		/// Perform the mail merge process.
		/// </summary>
		/// <param name="nameFormat">Optional, specify the format of the output PDF document names</param>
		public void mailMerge(String nameFormat = "PDF Mail Merge {MAILMERGE_ID}")
		{
			// Check preconditions
			if (this.getTemplatePath() != null && this.getFillableData() != null && this.getDestinationPath() != null)
			{
				// Store the destination path into a string for easy access
				String filepath = this.getDestinationPath();

				// Get the columns of the data table
				String[] fields = new String[fillableData.Columns.Count];

				DataColumnCollection cols = fillableData.Columns;
				int i = 0;
				foreach (DataColumn column in cols)
				{
					fields[i] = column.ToString();
					i++;
				}

				// Iterate rows of data table, generate PDF
				i = 0;
				int mailMergeIdPadding = this.fillableData.Rows.Count.ToString().Length;
				foreach (DataRow row in this.fillableData.Rows)
				{
					// Load the PDF template into the pdfReader - must do this for each one
					//try
					//{
					PdfReader pdfReader = new PdfReader(this.getTemplatePath());
					//}
					//catch()
					//{

					//}

					// Make file name
					string filename = nameFormat;
					foreach (char c in Path.GetInvalidFileNameChars())
					{
						filename = filename.Replace(c, '_');
					}

					if (filename.IndexOf("{") == -1)
					{
						// There is no unique identifier
						filename += " {MAILMERGE_ID}";
					}

					if (filename.LastIndexOf(".pdf") != filename.Length - 4)
					{
						// Ensure a PDF extension
						filename += ".pdf";
					}

					//Check for empty rows and replace the field name in the filename format
					bool emptyRow = true;
					foreach (String field in fields)
					{
						if (row[field].ToString().Length > 0)
						{
							emptyRow = false;
							filename = filename.Replace("{" + field + "}", row[field].ToString());
						}
					}
					if (emptyRow)
					{
						// Skip the rest of the processing, this is an empty row
						break;
					}

					filename = filename.Replace("{MAILMERGE_ID}", (i + 1).ToString().PadLeft(mailMergeIdPadding, '0'));

					// Create the output file
					FileMode mode = FileMode.CreateNew;
					if (File.Exists(filepath + filename))
					{
						mode = FileMode.Truncate;
					}
					PdfStamper pdfStamper = new PdfStamper(pdfReader, new FileStream(filepath + filename, mode));

					AcroFields pdfFormFields = pdfStamper.AcroFields;
					foreach (String field in fields)
					{
						// Only try to fill the field if the field has something in it
						if (row[field].ToString().Length > 0 && row[field].ToString().ToUpper() != "NULL")
						{

							pdfFormFields.SetField(field, row[field].ToString());
						}
					}

					pdfStamper.FormFlattening = false;
					pdfStamper.Close();

					i++;
				}
			}
			else
			{
				// TODO: Argument Exception is probably not best, class variables are not "arguments" per se
				if (this.getTemplatePath() == null)
				{
					throw new System.ArgumentException("Not all parameters set", "templatePath");
				}

				if (this.getDestinationPath() == null)
				{
					throw new System.ArgumentException("Not all parameters set", "fillableData");
				}

				if (this.getFillableData() == null)
				{
					throw new System.ArgumentException("Not all parameters set", "fillableData");
				}
			}
		}

		/// <summary>
		/// Get the path to the PDF Template
		/// </summary>
		/// <returns>Path to the PDF Template</returns>
		public String getTemplatePath()
		{
			return this.templatePath;
		}

		/// <summary>
		/// Set the path to the PDF Template
		/// </summary>
		/// <param name="path">Path to the PDF Template</param>
		public void setTemplatePath(String path)
		{
			if (File.Exists(path))
			{
				this.templatePath = path;
			}
		}

		/// <summary>
		/// Get the path to the output directory
		/// </summary>
		/// <returns>Path to the output directory</returns>
		public String getDestinationPath()
		{
			return this.destinationPath;
		}

		/// <summary>
		/// Set the path to the output directory
		/// </summary>
		/// <param name="path">Path to the output directory</param>
		public void setDestinationPath(String path)
		{
			if (Directory.Exists(path))
			{
				this.destinationPath = path;
			}
		}

		/// <summary>
		/// Get the data to be filled into the PDF template
		/// </summary>
		/// <returns>Data to be filled into the PDF template</returns>
		public DataTable getFillableData()
		{
			return this.fillableData;
		}

		/// <summary>
		/// Set the data to be filled into the PDF template
		/// </summary>
		/// <param name="data">Data to be filled into the PDF template</param>
		public void setFillableData(DataTable data)
		{
			this.fillableData = data;
		}
	}
}
