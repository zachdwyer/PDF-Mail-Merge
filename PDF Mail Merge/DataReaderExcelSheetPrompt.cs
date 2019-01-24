using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace PDF_Mail_Merge
{
	public partial class DataReaderExcelSheetPrompt : Form
	{
		public String SelectedSheet;

		public DataReaderExcelSheetPrompt()
		{
			InitializeComponent();
		}

		public DataReaderExcelSheetPrompt(String[] options)
		{
			InitializeComponent();

			foreach (String sheetName in options)
			{
				comboBox1.Items.Add(sheetName);
			}
		}

		private void button1_Click(object sender, EventArgs e)
		{
			if (comboBox1.SelectedItem != null)
			{
				SelectedSheet = comboBox1.SelectedItem.ToString();
				this.Close();
			}
			else
			{
				this.DialogResult = DialogResult.Cancel;
			}
		}


	}
}
