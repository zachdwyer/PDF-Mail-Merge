using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

namespace PDF_Mail_Merge
{
	public partial class Main : Form
	{
		private AboutBox about;
		private HelpFileNameFormat helpfnf;
		private MethodInvoker processInvoker;

		public Main()
		{
			InitializeComponent();
		}

		private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
		{
			about = new AboutBox();
			about.Show();
		}

		private void fileNameFormatToolStripMenuItem_Click(object sender, EventArgs e)
		{
			helpfnf = new HelpFileNameFormat();
			helpfnf.Show();
		}

		private void exitToolStripMenuItem_Click(object sender, EventArgs e)
		{
			this.Close();
		}

		private void buttonDataSource_Click(object sender, EventArgs e)
		{
			if (openFileDialog1.Filter.Length == 0)
			{
				// If the file filter is empty, populate it with dynamically generated list
				String filter = "";
				Type[] types = Assembly.GetExecutingAssembly().GetTypes();
				foreach (Type type in types)
				{
					// If this type is a subclass of DataReader, call its getFileFilter() function
					if (type.IsSubclassOf(typeof(DataReader)))
					{
						//Find all Constructors with no arguments (Type.EmptyTypes) and call them - with no arguments (null)
						DataReader fileReader = (DataReader)type.GetConstructor(Type.EmptyTypes).Invoke(null);
						filter += fileReader.getFileFilter() + "|";
					}
				}

				// Trim the trailing | char from the filter list
				openFileDialog1.Filter = filter.TrimEnd('|');
			}

			openFileDialog1.FileName = "";
			if (openFileDialog1.ShowDialog() == DialogResult.OK)
			{
				textBoxDataSource.Text = openFileDialog1.FileName;
			}
		}

		private void buttonTemplate_Click(object sender, EventArgs e)
		{
			openFileDialog2.FileName = "";
			if (openFileDialog2.ShowDialog() == DialogResult.OK)
			{
				textBoxTemplate.Text = openFileDialog2.FileName;
			}
		}

		private void buttonSavePath_Click(object sender, EventArgs e)
		{
			if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
			{
				textBoxSavePath.Text = folderBrowserDialog1.SelectedPath;
			}
		}

		private void buttonMerge_Click(object sender, EventArgs e)
		{
			processInvoker = new MethodInvoker(Merge_Process);
			processInvoker.BeginInvoke(null, null);
		}

		private void Merge_Process()
		{
			// Validate that the required fields are filled
			if (textBoxDataSource.Text.Trim().Length != 0
				&& textBoxTemplate.Text.Trim().Length != 0
				&& textBoxSavePath.Text.Trim().Length != 0)
			{
				// Default the document to an excel reader to avoid warnings below
				DataReader document = new DataReaderExcel();

				// Start going through all of the DataReader subclasses
				Type[] types = Assembly.GetExecutingAssembly().GetTypes();
				foreach (Type type in types)
				{
					// If this type is a subclass of DataReader, call its getFileFilterExtensions() function
					if (type.IsSubclassOf(typeof(DataReader)))
					{
						// Find all Constructors with no arguments (Type.EmptyTypes) and call them - with no arguments (null)
						DataReader tempDR = (DataReader)type.GetConstructor(Type.EmptyTypes).Invoke(null);

						// If the data file matches the extension for this particular DataReader subclass
						if (tempDR.getFileFilterExtensions().Contains(System.IO.Path.GetExtension(textBoxDataSource.Text.Trim())))
						{
							// Declare the document to be of this subclass and stop looking for subclasses
							document = (DataReader)type.GetConstructor(Type.EmptyTypes).Invoke(null);
							break;
						}
					}
				}

				DataTable fillableData = new DataTable();

				// If the file exists, and it parses correctly
				if (document.loadFile(textBoxDataSource.Text.Trim()))
				{
					// Get its dataset
					fillableData = document.getDataSet();

					// Establish filepath for PDF exports
					String filepath = textBoxSavePath.Text.Trim();
					if (filepath[filepath.Length - 1] != '\\')
					{
						//Ensure the trailing slash
						filepath += '\\';
					}
					foreach (char c in Path.GetInvalidPathChars())
					{
						filepath = filepath.Replace(c, '_');
					}

					// Make sure filepath exists
					if (!Directory.Exists(filepath))
					{
						Directory.CreateDirectory(filepath);
					}

					// Create a merge workhorse
					Merger mergeWorker = new Merger(textBoxTemplate.Text.Trim(), filepath, fillableData);

					//fillableData = mergeWorker.getFillableData();

					// Merge the data into a set of PDFs
					String fileNameFormat = textBoxFileNameFormat.Text.Trim();

					mergeWorker.mailMerge((fileNameFormat.Length > 0) ? fileNameFormat : null);
				}
				else
				{
					MessageBox.Show(textBoxDataSource.Text.Trim() + " either does not exist or cannot be parsed.", "Error");
				}
			}
			else
			{
				MessageBox.Show("You are missing some important fields.", "Error");
			}
		}
	}
}
