using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sharepoint
{
    class CreateTable
    {

        public static void CreateT()


            {



            List<string> ValueType1 = Extras.getcolumnnames();
            //string joined = string.Join(",", ValueType1);
            string output = "[" + string.Join("],[", ValueType1) + "]";

            String strConnection = Global.sqlcon;
            String excelConnString = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0\"", Global.filePath);

            using (OleDbConnection excelConnection = new OleDbConnection(excelConnString))
            {
                //Create OleDbCommand to fetch data from Excel 
                using (OleDbCommand cmd = new OleDbCommand("Select" + output + "from [Work Order Insights$]", excelConnection))
                //using (OleDbCommand cmd = new OleDbCommand("Select [Completed Fiscal Month],[MA],[Charter Business Flag],[Work Order Number],[Customer - Account Number],[Technician As Was - Technician Title],[Self Install - Rescue Flag],[Drop Bury Flag],[Primary Reason Code - Billing],[Technician As Was - Contracting Firm Name],[Technician As Was - Contractor / In-House],[Time 24 - WO Job Completed Time],[Fiber Node],[SVC Unit - Drop Hub Code],[Primary Resolution Code - Billing],[Positive Work Flag],[Work Order Job Class],[Completed Date],[Work Order Count] from [Work Order Insights$]", excelConnection))

                {
                    excelConnection.Open();
                    //FileUtil.WhoIsLocking(path);
                    SqlConnection SQLConnection = new SqlConnection();
                    SQLConnection.ConnectionString = Global.sqlcon;

                    DataTable dtSheet = excelConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    string sheetname;
                    sheetname = "";
                    foreach (DataRow drSheet in dtSheet.Rows)
                    {
                        if (drSheet["TABLE_NAME"].ToString().Contains("$"))
                        {
                            sheetname = drSheet["TABLE_NAME"].ToString();

                            //Load the DataTable with Sheet Data
                            OleDbCommand oconn = new OleDbCommand("select * from [" + sheetname + "]", excelConnection);
                            OleDbDataAdapter adp = new OleDbDataAdapter(oconn);
                            DataTable dt = new DataTable();
                            adp.Fill(dt);

                            //remove "$" from sheet name
                            sheetname = sheetname.Replace("$", "");

                            // Generate Create Table Script by using Header Column,
                            //It will drop the table if Exists and Recreate                  
                            string tableDDL = "";
                            tableDDL += "IF EXISTS (SELECT * FROM sys.objects WHERE object_id = ";
                            tableDDL += "OBJECT_ID(N'[dbo].[" + Global.filePath + "_" + sheetname + "]') AND type in (N'U'))";
                            tableDDL += "Drop Table [dbo].[" + Global.filePath + "_" + sheetname + "]";
                            tableDDL += "Create table [" + Global.filePath + "_" + sheetname + "]";
                            tableDDL += "(";
                            for (int i = 0; i < dt.Columns.Count; i++)
                            {
                                if (i != dt.Columns.Count - 1)
                                    tableDDL += "[" + dt.Columns[i].ColumnName + "] " + "NVarchar(max)" + ",";
                                else
                                    tableDDL += "[" + dt.Columns[i].ColumnName + "] " + "NVarchar(max)";
                            }
                            tableDDL += ")";

                            SQLConnection.Open();
                            SqlCommand SQLCmd = new SqlCommand(tableDDL, SQLConnection);
                            SQLCmd.ExecuteNonQuery();

                            ////Load the data from DataTable to SQL Server Table.
                            //SqlBulkCopy blk = new SqlBulkCopy(SQLConnection);
                            //blk.DestinationTableName = "[" + filename + "_" + sheetname + "]";
                            //blk.WriteToServer(dt);
                            //SQLConnection.Close();
                        }

                    }


                    //using (OleDbDataReader dReader = cmd.ExecuteReader())
                    //{
                    //    using (SqlBulkCopy sqlBulk = new SqlBulkCopy(strConnection))
                    //    {
                    //        //Give your Destination table name 
                    //        sqlBulk.DestinationTableName = "DropData";
                    //        sqlBulk.WriteToServer(dReader);
                    //    }
                    //}






                }
            }







            ////the datetime and Log folder will be used for error log file in case error occured
            //string datetime = DateTime.Now.ToString("yyyyMMddHHmmss");
            //string LogFolder = @"C:\Log\";
            //try
            //{
            //    //Provide the Source Folder path where excel files are present
            //    String FolderPath = @"C:\Reports\";
            //    //Provide the Database Name 
            //    //string DatabaseName = "Test";
            //    //Provide the SQL Server Name 
            //    //string SQLServerName = "(local)";


            //    //Create Connection to SQL Server Database 
            //    SqlConnection SQLConnection = new SqlConnection();
            //    SQLConnection.ConnectionString = Global.sqlcon;

            //    var directory = new DirectoryInfo(FolderPath);
            //    FileInfo[] files = directory.GetFiles();

            //    //Declare and initilize variables
            //    string fileFullPath = "";

            //    //Get one Book(Excel file at a time)
            //    foreach (FileInfo file in files)
            //    {
            //        string filename = "";
            //        fileFullPath = FolderPath + "\\" + file.Name;
            //        filename = file.Name.Replace(".xlsx", "");

            //        //Create Excel Connection
            //        string ConStr;
            //        string HDR;
            //        HDR = "YES";
            //        ConStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileFullPath
            //            + ";Extended Properties=\"Excel 12.0;HDR=" + HDR + ";IMEX=1\"";
            //        OleDbConnection cnn = new OleDbConnection(ConStr);


            //        //Get Sheet Name
            //        cnn.Open();
            //        DataTable dtSheet = cnn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            //        string sheetname;
            //        sheetname = "";
            //        foreach (DataRow drSheet in dtSheet.Rows)
            //        {
            //            if (drSheet["TABLE_NAME"].ToString().Contains("$"))
            //            {
            //                sheetname = drSheet["TABLE_NAME"].ToString();

            //                //Load the DataTable with Sheet Data
            //                OleDbCommand oconn = new OleDbCommand("select * from [" + sheetname + "]", cnn);
            //                OleDbDataAdapter adp = new OleDbDataAdapter(oconn);
            //                DataTable dt = new DataTable();
            //                adp.Fill(dt);

            //                //remove "$" from sheet name
            //                sheetname = sheetname.Replace("$", "");

            //                // Generate Create Table Script by using Header Column,
            //                //It will drop the table if Exists and Recreate                  
            //                string tableDDL = "";
            //                tableDDL += "IF EXISTS (SELECT * FROM sys.objects WHERE object_id = ";
            //                tableDDL += "OBJECT_ID(N'[dbo].[" + filename + "_" + sheetname + "]') AND type in (N'U'))";
            //                tableDDL += "Drop Table [dbo].[" + filename + "_" + sheetname + "]";
            //                tableDDL += "Create table [" + filename + "_" + sheetname + "]";
            //                tableDDL += "(";
            //                for (int i = 0; i < dt.Columns.Count; i++)
            //                {
            //                    if (i != dt.Columns.Count - 1)
            //                        tableDDL += "[" + dt.Columns[i].ColumnName + "] " + "NVarchar(max)" + ",";
            //                    else
            //                        tableDDL += "[" + dt.Columns[i].ColumnName + "] " + "NVarchar(max)";
            //                }
            //                tableDDL += ")";

            //                SQLConnection.Open();
            //                SqlCommand SQLCmd = new SqlCommand(tableDDL, SQLConnection);
            //                SQLCmd.ExecuteNonQuery();

            //                //Load the data from DataTable to SQL Server Table.
            //                SqlBulkCopy blk = new SqlBulkCopy(SQLConnection);
            //                blk.DestinationTableName = "[" + filename + "_" + sheetname + "]";
            //                blk.WriteToServer(dt);
            //                SQLConnection.Close();
            //            }

            //        }
            //    }
            //}
            //catch (Exception exception)
            //{
            //    // Create Log File for Errors
            //    using (StreamWriter sw = File.CreateText(LogFolder
            //        + "\\" + "ErrorLog_" + datetime + ".log"))
            //    {
            //        sw.WriteLine(exception.ToString());

            //    }

            //}




        }








    }
}
