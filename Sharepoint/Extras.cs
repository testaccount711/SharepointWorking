using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sharepoint
{
    class Extras
    {
        static public void CopyFileAccess()
        {
            var directory = new DirectoryInfo(@"\\tamp20pvfiler09\SHARE1\cfdivix1\Public\Finance\Business Planning\Daily\Reporting Tools\Proj\Reports\BIDownloads\");
            var myFile = (from f in directory.GetFiles()
                          orderby f.CreationTime descending
                          select f).First();



            string path = @"\\tamp20pvfiler09\SHARE1\cfdivix1\Public\Finance\Business Planning\Daily\Reporting Tools\Proj\Reports\BIDownloads\" + myFile.Name;
            string path2 = @"\\tamp20pvfiler09\SHARE1\cfdivix1\Public\Finance\Business Planning\Daily\Reporting Tools\Proj\Reports\BIDownloads\New" + myFile.Name;

            //        using (var inputFile = new FileStream(
            //path,
            //FileMode.Open,
            //FileAccess.Read,
            //FileShare.ReadWrite))
            //        {
            //            using (var outputFile = new FileStream(path2, FileMode.Create))
            //            {
            //                var buffer = new byte[0x10000];
            //                int bytes;

            //                while ((bytes = inputFile.Read(buffer, 0, buffer.Length)) > 0)
            //                {
            //                    outputFile.Write(buffer, 0, bytes);
            //                }
            //            }
            //        }

            FileStream inf = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            FileStream outf = new FileStream(path2, FileMode.Create);
            int a;
            while ((a = inf.ReadByte()) != -1)
            {
                outf.WriteByte((byte)a);
            }
            inf.Close();
            inf.Dispose();
            outf.Close();
            outf.Dispose();











        }


        public static List<String> ReadSpecificTableColumns(string filePath, string sheetName)
        {
            var columnList = new List<string>();
            try
            {
                var excelConnection = new OleDbConnection(String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0\"", filePath));
                excelConnection.Open();

                DataSet dataSet = new DataSet();
                List<string> listColumn = new List<string>();
                foreach (DataTable table in dataSet.Tables)
                {
                    foreach (DataColumn column in table.Columns)
                    {

                        ///listColumn.Add(column["Column_name"].ToString());
                        Console.WriteLine(column.ColumnName);
                        Console.ReadKey();
                    }
                }



                //List<string> listColumn = new List<string>();
                //foreach (DataRow row in dt.Rows)
                //{
                //    listColumn.Add(row["Column_name"].ToString());
                //    Console.WriteLine(listColumn.);
                //}


                ///columnList.AddRange(from DataRow column in columns.Rows select column["Column_name"].ToString());


                //excelConnection.Close();


            }
            catch (Exception exception)
            {
                Console.WriteLine(exception.Message);
            }

            return columnList;
        }

        public static List<string> getcolumnnames()
        //public static void getcolumnnames()
        {
            var excelWorkbook = new XLWorkbook(Global.filePath);


            IXLWorksheet workSheet = excelWorkbook.Worksheet(1);
            //var range = workSheet.FirstRowUsed();

            //var table = range.AsTable();
            string B = workSheet.Name;

            var tableList = workSheet.FirstRowUsed()
                     .CellsUsed()
                         .Select(c => c.Value)
                             .ToList();

            List<string> VorfahrCK = new List<string>();
            /// Console.WriteLine(range.RangeAddress);
            foreach (string item in tableList)
            {
                //Console.WriteLine(item);
                //Console.Write(item);
                VorfahrCK.Add(item);
            }



            return VorfahrCK;
        
            //VorfahrCK.ForEach(Console.WriteLine);
            //Console.ReadKey();

        }


        public static void ReadFile()
        {

        
            var workbook = new XLWorkbook(Global.filePath);
            var ws1 = workbook.Worksheet(1);
            var row = ws1.Row(1);
            row.Delete();
            workbook.SaveAs(Global.filePath);





        }











    }



}
