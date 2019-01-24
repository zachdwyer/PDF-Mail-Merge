using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;

namespace PDF_Mail_Merge
{
	abstract class DataReader
	{
		/// <summary>
		/// Stores the dataset after parsing.
		/// </summary>
		public DataTable fillableDataSet;

		/// <summary>
		/// Stores the path of the file to be parsed.
		/// </summary>
		public String filePath;

		/// <summary>
		/// Default constructor
		/// </summary>
		public DataReader()
			: this(null, null)
		{
		}

		/// <summary>
		/// Construct the object and specify a file path for a file to be parsed.
		/// </summary>
		/// <param name="filePath">Path of the file to be parsed.</param>
		public DataReader(String filePath)
			: this(filePath, null)
		{
		}

		/// <summary>
		/// Construct the object and specify the dataset to work with.
		/// </summary>
		/// <param name="dataSet">The dataset.</param>
		public DataReader(DataTable dataSet)
			: this(null, dataSet)
		{
		}

		/// <summary>
		/// Contruct the object and specify both a file path of a file to be parsed and a dataset to work with.
		/// </summary>
		/// <param name="filePath">Path of the file to be parsed.</param>
		/// <param name="dataSet">The dataset.</param>
		public DataReader(String filePath, DataTable dataSet)
		{
			if (dataSet == null)
			{
				dataSet = new DataTable();
			}

			setFilePath(filePath);
			setDataSet(dataSet);
		}

		/// <summary>
		/// Sets the filePath variable and parses the file.
		/// </summary>
		/// <param name="path">The file path.</param>
		/// <returns>Success of loading the file and parsing it.</returns>
		public bool loadFile(String path)
		{
			if (File.Exists(path))
			{
				setFilePath(path);
				return parseFile();
			}

			return false;
		}

		/// <summary>
		/// Initiate a parse of the loaded file.
		/// </summary>
		/// <returns>Success of the parse operation.</returns>
		abstract public bool parseFile();

		/// <summary>
		/// Get the parsed dataset.
		/// </summary>
		/// <returns>The parsed dataset.</returns>
		public DataTable getDataSet()
		{
			return this.fillableDataSet;
		}

		/// <summary>
		/// Sets the dataset.
		/// </summary>
		/// <param name="dataSet">The dataset.</param>
		public void setDataSet(DataTable dataSet)
		{
			this.fillableDataSet = dataSet;
		}

		/// <summary>
		/// Get the path of the file to be parsed.
		/// </summary>
		/// <returns>Path of the file to be parsed.</returns>
		public String getFilePath()
		{
			return this.filePath;
		}

		/// <summary>
		/// Set the path of the file to be parsed.
		/// </summary>
		/// <param name="path">Path of the file to be parsed.</param>
		public void setFilePath(String path)
		{
			if (File.Exists(path))
			{
				this.filePath = path;
			}
		}

		/// <summary>
		/// Get the file type filter for a browse for file button.
		/// </summary>
		abstract public String getFileFilter();

		/// <summary>
		/// Get the file extension list that this DataReader parses.
		/// </summary>
		abstract public String[] getFileFilterExtensions();
	}
}
