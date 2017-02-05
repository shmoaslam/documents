using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using Excel;
using System.Data;
using System.Text;

namespace UnitTestProject
{

    public static class ExtensionClass
    {
        public static string DataTableToCSV(this DataTable datatable, char seperator)
        {
            StringBuilder sb = new StringBuilder();
            //for (int i = 0; i < datatable.Columns.Count; i++)
            //{
            //    sb.Append(datatable.Columns[i]);
            //    if (i < datatable.Columns.Count - 1)
            //        sb.Append(seperator);
            //}
            //sb.AppendLine();
            foreach (DataRow dr in datatable.Rows)
            {
                for (int i = 0; i < datatable.Columns.Count; i++)
                {
                    sb.Append(dr[i].ToString());

                    if (i < datatable.Columns.Count - 1)
                        sb.Append(seperator);
                }
                sb.AppendLine();
            }
            return sb.ToString();
        }
    }


    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            string filesPath = @"C:\Users\shmoa\Downloads\MadhavTask\";

            if (!Directory.Exists(filesPath))
                return;

            var excelFiles = Directory.GetFiles(filesPath, "*.xlsx");

            foreach (var excelfile in excelFiles)
            {
                var excelFilePath = Path.Combine(filesPath, excelfile);

                if (!File.Exists(excelFilePath))
                    continue;

                var excelFileNameWithOutExtension = Path.GetFileNameWithoutExtension(excelFilePath);

                IExcelDataReader excelReader;
                try
                {
                    FileStream stream = File.Open(excelFilePath, FileMode.Open, FileAccess.Read);

                    if (stream == null) continue;

                    excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);

                    if (excelReader == null) continue;

                    DataSet sheetsData = excelReader.AsDataSet();

                    if (sheetsData == null) continue;

                    if (sheetsData.Tables.Count > 0)
                    {
                        for (int i = 0; i < sheetsData.Tables.Count; i++)
                        {
                            var table = (DataTable)sheetsData.Tables[i];

                            if (table == null) continue;

                            var tableName = table.TableName;

                            var subFolderPathForCVSTabs = Path.Combine(filesPath, tableName);

                            if (!Directory.Exists(subFolderPathForCVSTabs))
                                Directory.CreateDirectory(subFolderPathForCVSTabs);

                            var csvResult = table.DataTableToCSV(',');

                            if (csvResult == null) continue;

                            var fileName = string.Format("{0}_{1}.csv", excelFileNameWithOutExtension, tableName);

                            File.WriteAllText(subFolderPathForCVSTabs  + @"\"+ fileName , csvResult.ToString());
                        }
                    }
                    excelReader.Close();
                }
                catch (Exception ex)
                {
                    throw;
                }
            }


            //// Reading from a binary Excel file ('97-2003 format; *.xls)
            //IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(stream);

            // Reading from a OpenXml Excel file (2007 format; *.xlsx)

            // DataSet - The result of each spreadsheet will be created in the result.Tables
            // Free resources (IExcelDataReader is IDisposable)

        }
    }
}
