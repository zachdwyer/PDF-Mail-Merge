using System;
using System.IO;
using System.Globalization;
using System.Reflection;
using System.Resources;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PDF_Mail_Merge;
using System.Data;

namespace PDF_Mail_Merge_Tests
{
	/// <summary>
	///This is a test class for DataReaderExcelTest and is intended
	///to contain all DataReaderExcelTest Unit Tests
	///</summary>
	[TestClass()]
	public class DataReaderExcelTest
	{
		/// <summary>
		/// The path to a document containing test data
		/// </summary>
		String testDataSource;

		private TestContext testContextInstance;

		/// <summary>
		///Gets or sets the test context which provides
		///information about and functionality for the current test run.
		///</summary>
		public TestContext TestContext
		{
			get
			{
				return testContextInstance;
			}
			set
			{
				testContextInstance = value;
			}
		}

		#region Additional test attributes
		// 
		//You can use the following additional attributes as you write your tests:
		//
		//Use ClassInitialize to run code before running the first test in the class
		//[ClassInitialize()]
		//public static void MyClassInitialize(TestContext testContext)
		//{
		//}
		//
		//Use ClassCleanup to run code after all tests in a class have run
		//[ClassCleanup()]
		//public static void MyClassCleanup()
		//{
		//}
		#endregion

		// TestInitialize to run code before running each test
		[TestInitialize()]
		public void MyTestInitialize()
		{
			// Establish a test document path
			String uniqueFileName = Path.GetTempFileName();
			this.testDataSource = uniqueFileName + ".xlsx";
			File.Move(uniqueFileName, this.testDataSource);

			// Write test document from assembly resources to path
			File.WriteAllBytes(this.testDataSource, Resources.TestDocument);
		}
		
		// TestCleanup to run code after each test has run
		[TestCleanup()]
		public void MyTestCleanup()
		{
			// Cleanup the test document
			if (File.Exists(this.testDataSource))
			{
				File.Delete(this.testDataSource);
			}
		}

		/// <summary>
		///A test for parseFile
		///</summary>
		[TestMethod()]
		public void parseFileTest()
		{
			DataReaderExcel target = new DataReaderExcel();
			target.setFilePath(this.testDataSource);
			
			bool expected = true;
			bool actual;
			actual = target.parseFile("Sheet1");

			// Check that it parsed correctly
			Assert.AreEqual<bool>(expected, actual);
			
			// Check parsed
			DataTable results = target.getDataSet();
			Assert.IsNotNull(results);

			// Check column names
			Assert.AreEqual<String>(results.Columns[0].ColumnName, "FIRST_NAME");
			Assert.AreEqual<String>(results.Columns[1].ColumnName, "LAST_NAME");

			// Check strings
			Assert.AreEqual<String>((String)results.Rows[0][0], "Reed");
			Assert.AreEqual<String>((String)results.Rows[0][1], "Richards");

			Assert.AreEqual<String>((String)results.Rows[7][0], "Thor");

			// Check Empty String
			Assert.AreEqual<String>((String)results.Rows[7][1], "");
		}

		/// <summary>
		///A test for formatData
		///</summary>
		[TestMethod()]
		[DeploymentItem("PDF Mail Merge.exe")]
		public void formatDataTest()
		{
			// Parse Test Document
			DataReaderExcel target = new DataReaderExcel();
			target.setFilePath(this.testDataSource);
			target.parseFile("Sheet1");
			DataTable results = target.getDataSet();

			// Check date format
			Object obj = results.Rows[4][10];
			DateTime date = new DateTime(2012, 5, 29);
			Assert.AreEqual<String>((String)results.Rows[4][10], date.ToShortDateString());

			date = new DateTime(2012, 7, 4);
			Assert.AreEqual<String>((String)results.Rows[4][10], date.ToShortDateString());
		}
	}
}
