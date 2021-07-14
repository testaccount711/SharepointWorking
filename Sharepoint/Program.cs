using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Microsoft.SharePoint.Client;
using System.Net;
using File = System.IO.File;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Data;
using System.Threading;
using System.Runtime.InteropServices;
using OfficeOpenXml;
using System.Net.Mail;

namespace Sharepoint
{
    class Program
    {
        static void Main(string[] args)
        {
            ///////////////// Transfer data to SQL Server From Sharepoint Site
            //SharepointDownload();
           //TruncTable();
            //UploadData();
             UploadDataBI();
            //Sharepoint.Extras.CopyFileAccess();

            ////TruncTable();    -- need to add truncate capability to svc account
            //UploadData();

            //////////////////////////// Assembling Reports Below
            //assemble();
            //assembleknight();
            //assembleKablelink();
            //assembleAeon();
            //assembleJaguar();
            //////////////////////////// Emailing Recipients Below

            //MailFileInhouse();
            //MailFileAeon();
            //MailFileKablelink();
            //MailFileJag();

            //Extras.ReadFile();
            //CreateTable.CreateT();


        }


        static void SharepointDownload()
        {

            using (ClientContext ctx = new ClientContext("https://sharepoint.charter.com/ops/FO-Florida/Reporting/"))
            {

                ctx.Credentials = new NetworkCredential("P2162735", "Harrydog34", "CHTR");


                FileCollection files = ctx.Web.GetFolderByServerRelativeUrl("https://sharepoint.charter.com/ops/FO-Florida/Reporting/Florida%20Hotspots").Files;
                ctx.Load(files);
                if (ctx.HasPendingRequest)
                {
                    ctx.ExecuteQuery();
                }
                Console.WriteLine("Connecting to Sharepoint");

                System.IO.DirectoryInfo di = new DirectoryInfo(@"\\tamp20pvfiler09\SHARE1\cfdivix1\Public\Finance\Business Planning\Daily\Reporting Tools\Proj\HotSpot\");

                foreach (FileInfo file in di.GetFiles())
                {
                    file.Delete();
                }

                Console.WriteLine("Deleting Old Files in Target Directory");

                foreach (Microsoft.SharePoint.Client.File file in files)
                {

                    if (file.TimeCreated >= DateTime.Now.AddHours(-8))

                    {
                        using (FileInformation fileInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(ctx, file.ServerRelativeUrl))
                        {
                            ctx.ExecuteQuery();

                            var filePath = @"\\tamp20pvfiler09\SHARE1\cfdivix1\Public\Finance\Business Planning\Daily\Reporting Tools\Proj\HotSpot\" + file.Name;
                            var fileName = Path.Combine(filePath, Path.GetFileName(file.ServerRelativeUrl));
                            FileStream fileStream = new FileStream(filePath, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite);
                            fileInfo.Stream.CopyTo(fileStream);
                            fileStream.Close();
                            ctx.Dispose();
                            Console.WriteLine("Downloading New File");
                            Thread.Sleep(3000);
                        }
                    }
                    
                



                }

            }
        }


        public static void GetNewest()
        {


            var directory = new DirectoryInfo(@"\\tamp20pvfiler09\SHARE1\cfdivix1\Public\Finance\Business Planning\Daily\Reporting Tools\Proj\HotSpot\");
            var myFile = (from f in directory.GetFiles()
                          orderby f.CreationTime ascending
                          select f).First();
            Console.WriteLine(myFile);
            Console.ReadLine();

            //DirectoryInfo info = new DirectoryInfo(@"\\tamp20pvfiler09\SHARE1\cfdivix1\Public\Finance\Business Planning\Daily\Reporting Tools\Proj\HotSpot\");
            //FileInfo[] files = info.GetFiles().OrderBy(p => p.CreationTime).ToArray();
            //foreach (FileInfo file in files)
            //{
            //                Console.WriteLine(file);
            //                Console.ReadLine();
            //}

        }



        //Console.WriteLine(file.ToString());
        //Console.ReadLine();


        static void TruncTable()
        {

            SqlConnection con = new SqlConnection(Global.sqlcon);
            con.Open();
            SqlCommand cmd = new SqlCommand("TRUNCATE TABLE Archer", con);
            cmd.ExecuteNonQuery();
            con.Close();
            Console.WriteLine("Truncating Target Table");
            Thread.Sleep(5000);
        }

        static void UpdateTable()
        {

            SqlConnection con = new SqlConnection(Global.sqlcon);
            con.Open();
            SqlCommand cmdd = new SqlCommand("ATG1682Update", con);
            cmdd.CommandType = CommandType.StoredProcedure;
            cmdd.ExecuteNonQuery();
            con.Close();
            //Console.WriteLine("All Done");
            //Console.ReadKey();

        }


        private static void UploadData()
        {


            var directory = new DirectoryInfo(@"\\tamp20pvfiler09\SHARE1\cfdivix1\Public\Finance\Business Planning\Daily\Reporting Tools\Proj\HotSpot\");
            var myFile = (from f in directory.GetFiles()
                          orderby f.CreationTime descending
                          select f).First();



            string path = @"\\tamp20pvfiler09\SHARE1\cfdivix1\Public\Finance\Business Planning\Daily\Reporting Tools\Proj\HotSpot\" + myFile.Name;
            //Console.WriteLine(path);
            //Console.ReadKey();



            try
            {
                //string aFileName;
                //GetFile(out aFileName);
                String strConnection = Global.sqlcon;
                String excelConnString = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0\"", path);
                Console.WriteLine("Uploading New Data into Target Directory in SQL Server");
                using (OleDbConnection excelConnection = new OleDbConnection(excelConnString))
                {
                    //Create OleDbCommand to fetch data from Excel 
                    using (OleDbCommand cmd = new OleDbCommand("Select [Tracking ID], [Market],[Date BC Install Completed],[Date Original AP First Live],[AP Status], [Last Exception Update],[Last Exception Type],[Biller Account Name],[Biller Account Number],[Street Address 1],	[City],	[State], [Zip Code],[Work Order Number],[Technician Name],[Technician Number] from [Archer Search Report$]", excelConnection))
                    {
                        excelConnection.Open();
                        using (OleDbDataReader dReader = cmd.ExecuteReader())
                        {
                            using (SqlBulkCopy sqlBulk = new SqlBulkCopy(strConnection))
                            {
                                //Give your Destination table name 
                                sqlBulk.DestinationTableName = "Archer";
                                sqlBulk.WriteToServer(dReader);
                            }
                        }
                    }
                }


            }

            catch (Exception ex)
            {

                Console.WriteLine(string.Format("Message: {0}<br />{1}StackTrace :{2}{1}Date :{3}{1}-----------------------------------------------------------------------------{1}", ex.Message, Environment.NewLine, ex.StackTrace, DateTime.Now.ToString()));
                Console.ReadKey();
            }



        }

        private static void UploadDataBI()
        {


            var directory = new DirectoryInfo(@"\\tamp20pvfiler09\SHARE1\cfdivix1\Public\Finance\Business Planning\Daily\Reporting Tools\Proj\Reports\BIDownloads\");
            var myFile = (from f in directory.GetFiles()
                          orderby f.CreationTime descending
                          select f).First();



            string path = @"\\tamp20pvfiler09\SHARE1\cfdivix1\Public\Finance\Business Planning\Daily\Reporting Tools\Proj\Reports\BIDownloads\" + myFile.Name;
            //Console.WriteLine(path);
            //Console.ReadKey();

            //FileUtil.WhoIsLocking(path);
            //Console.WriteLine(IsFileinUse(path));
            //Console.ReadKey();

            //Console.WriteLine(Extras.ReadSpecificTableColumns(path, "Sheet1"));
            //Console.ReadKey();
            //List<string> ValueType1 = Extras.getcolumnnames();--
            //string joined = string.Join(",", ValueType1); -- 
            //string output = "[" + string.Join("],[", ValueType1) + "]";
            //File.WriteAllText(Global.filePath2, output);  --

            //if (IsFileinUse(path) == false)
            //{


            //    Thread.Sleep(2000);
            try
                {
                    //string aFileName;
                    //GetFile(out aFileName);
                    String strConnection = Global.sqlcon;
                    String excelConnString = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0\"", Global.filePath);
                    Console.WriteLine("Uploading BI Data into Target Directory in SQL Server");

                    

                    using (OleDbConnection excelConnection = new OleDbConnection(excelConnString))
                    {
                        //Create OleDbCommand to fetch data from Excel 
                        using (OleDbCommand cmd = new OleDbCommand("Select [Completed Fiscal Month],[MA],[Completed Date],[Account Tenure - Individual Months],[Primary Reason Code (Billing) SRO ICOMS],[Technician -  Locator Code],[WO Job Schedule Category Code],[WorkZone Code],[Bulk Flag],[Technician As Was - Contractor / In-House],[Customer - Account Number],[Appointment Window],[Chargeable Fee Applied Flag],[Chargeable Flag],[Dwelling Type - Category - As-Is],[Job Sequence],[MA-Hub-Node],[Meter Eligible Flag],[Meter Test Result Indicator - CPE],[Meter Test Result Indicator - Ground Block],[Meter Test Result Indicator - Pressure Test],[Meter Test Result Indicator - Tap],[Multipeat Flag],[Multipeat Initial WO Flag],[Multipeat Sequence Number],[Not Done - Verification Code],[Not Done - Verification Flag],[Not Done Flag],[OTA],[Previous Technician Name],[Primary Reason Product Category],[Primary Reason Product Sub Category],[Reason Code Desc 2],[Reason Code Desc 3],[Reason Code Desc 4],[Refer Drop Bury Flag],[Refer MDU Post Wire Flag],[Refer to Construction Flag],[Refer to Maintenance Flag],[Primary Resolution Code Category],[Resolution Code Desc 2],[Resolution Code Desc 3],[Resolution Code Desc 4],[Resolution Code Desc 5],[Resolution Code Desc 6],[Self Install - Rescue Flag],[SVC Unit - Dev Number],[SVC Unit - FTTP Flag],[SVC Unit - Latitude],[SVC Unit - Longitude],[SVC Unit - Original Created Date],[SVC Unit - Serviceability - Date],[SVC Unit - Zip Code],[SVC Unit - Zip Code 4],[Tech Title Group],[Technician As Was - Full Name],[Technician As Was - HR Number/Empl ID],[Technician As Was - Manager],[Technician As Was - Supervisor],[TR Tech Title],[Unnecessary Truck Roll Category Desc],[WO - Activity Category],[WO - Activity Group],[WO - Activity Outcome],[WO - Activity Preventable Indicator],[WO - Activity Rework],[WO - Activity Type],[WO Job Comments],[WO Job Resched Cnt],[WO Job Start DTTM],[WO Job Stop DTTM],[WO Job Technician Comments - Rework Completed],[Work Order Number],[30 Day Repeat Flag - Initial WO],[Technician Id - WO Job Tech],[30 Day Repeat Flag - Repeat WO],[Routing Area Code],[Company Code],[Dwelling Type - Description - As-Is],[Primary Reason Code Desc],[Primary Resolution Code Desc],[WO Account Type Segment],[Entered Date Time],[Primary Reason Code - Billing],[Entered Date],[Work Order Number - Rework Completed],[Work Order Count],[Job Minutes - Avg],[Tech Job Minutes],[Quota Minutes],[Quota Points],[HHC - Compliant WO Cnt],[HHC - Device Compliance %],[HHC - Device Compliance Cnt],[HHC - Device Tested Cnt],[HHC - Eligible WO Cnt],[HHC - Overall Compliance %],[HHC - WO Test Completed Cnt],[HHC - WO Usage %] from [Work Order Insights$]", excelConnection))
                    //using (OleDbCommand cmd = new OleDbCommand("Select [Completed Fiscal Month],[MA],[Charter Business Flag],[Work Order Number],[Customer - Account Number],[Technician As Was - Technician Title],[Self Install - Rescue Flag],[Drop Bury Flag],[Primary Reason Code - Billing],[Technician As Was - Contracting Firm Name],[Technician As Was - Contractor / In-House],[Time 24 - WO Job Completed Time],[Fiber Node],[SVC Unit - Drop Hub Code],[Primary Resolution Code - Billing],[Positive Work Flag],[Work Order Job Class],[Completed Date],[Work Order Count] from [Work Order Insights$]", excelConnection))

                    {
                            excelConnection.Open();
                            //FileUtil.WhoIsLocking(path);
                            using (OleDbDataReader dReader = cmd.ExecuteReader())
                            {
                                using (SqlBulkCopy sqlBulk = new SqlBulkCopy(strConnection))
                                {
                                    //Give your Destination table name 
                                    sqlBulk.DestinationTableName = "WOI";
                                    sqlBulk.BulkCopyTimeout = 0;
                                    sqlBulk.WriteToServer(dReader);
                                }
                            }
                        }
                    }


                }

                catch (Exception ex)
                {

                    Console.WriteLine(string.Format("Message: {0}<br />{1}StackTrace :{2}{1}Date :{3}{1}-----------------------------------------------------------------------------{1}", ex.Message, Environment.NewLine, ex.StackTrace, DateTime.Now.ToString()));
                    Console.ReadKey();
                }

            //}
            //else

            //{

            //    Console.WriteLine("File is in use exiting......");
            //    Console.ReadKey();
            //    Environment.Exit(0);



            //}

        }









        static void assemble()
        {


            string sql = "Select Count(*) from Archer";
            var con1 = new SqlConnection(Global.sqlcon);

            SqlCommand comd = new SqlCommand(sql, con1);
            con1.Open();
            int ct = Convert.ToInt32(comd.ExecuteScalar());





            Console.WriteLine(ct.ToString() + " :Records");
            //Console.ReadKey();




            if (ct > 1)
            {
                TimeSpan stop;
                TimeSpan start = new TimeSpan(DateTime.Now.Ticks);


                DataTable dt = new DataTable();
                using (var con = new SqlConnection(Global.sqlcon))
                using (var cmd = new SqlCommand("ArcherSP", con))

                using (var da = new SqlDataAdapter(cmd))
                {
                    Console.WriteLine("Running Stored Procedure");
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandTimeout = 600;

                    da.Fill(dt);
                }




                string timestamp = DateTime.Now.ToString("MM-dd-yyyy");
                const string templatePath = @"\\tamp20pvfiler09\SHARE1\cfdivix1\Public\Finance\Business Planning\Daily\Reporting Tools\Proj\HotSpot\Template\template.xlsx"; // the path of the template
                string resultPath = @"\\tamp20pvfiler09\SHARE1\cfdivix1\Public\Finance\Business Planning\Daily\Reporting Tools\Proj\HotSpot\Reports\SMB Florida Hotspot Activations Exceptions" + " " + timestamp + ".xlsx"; // the path of our result
                File.Copy(templatePath, resultPath, true);
                var templateFile = new FileInfo(templatePath);
                var resultsFile = new FileInfo(resultPath);

                if (resultsFile.Exists)
                    resultsFile.Delete();



                using (var pck = new ExcelPackage(new FileInfo(resultPath), new FileInfo(templatePath)))

                {
                    Console.WriteLine("Assembling Report");
                    var ws = pck.Workbook.Worksheets["Data"];
                    var a = ws.Cells["A1"].Value;
                    int i = 2;
                    ws.InsertRow(i, dt.Rows.Count);
                    ws.Cells["A1"].LoadFromDataTable(dt, true);
                    ws.DeleteRow(dt.Rows.Count + 2);
                    pck.Save();
                    stop = new TimeSpan(DateTime.Now.Ticks);
                    GC.Collect();
                    GC.WaitForPendingFinalizers();


                }

            }

            else
            {

                Console.WriteLine("Row Count Less than 1 Exiting");
                Console.ReadKey();
                System.Environment.Exit(1);


            }



        }






        static void assembleknight()
        {

            Thread.Sleep(5000);

            TimeSpan stop;
            TimeSpan start = new TimeSpan(DateTime.Now.Ticks);


            DataTable dt = new DataTable();
            using (var con = new SqlConnection(Global.sqlcon))
            using (var cmd = new SqlCommand("ArcherSPKnight", con))

            using (var da = new SqlDataAdapter(cmd))
            {
                Console.WriteLine("Running Stored Procedure Knight Enterprises");
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandTimeout = 600;

                da.Fill(dt);
            }




            string timestamp = DateTime.Now.ToString("MM-dd-yyyy");
            const string templatePath = @"\\tamp20pvfiler09\SHARE1\cfdivix1\Public\Finance\Business Planning\Daily\Reporting Tools\Proj\HotSpot\Template\template.xlsx"; // the path of the template
            string resultPath = @"\\tamp20pvfiler09\SHARE1\cfdivix1\Public\Finance\Business Planning\Daily\Reporting Tools\Proj\HotSpot\Reports\Knight\SMB Florida Hotspot Activations Exceptions Knight" + " " + timestamp + ".xlsx"; // the path of our result
            File.Copy(templatePath, resultPath, true);
            var templateFile = new FileInfo(templatePath);
            var resultsFile = new FileInfo(resultPath);

            if (resultsFile.Exists)
                resultsFile.Delete();



            using (var pck = new ExcelPackage(new FileInfo(resultPath), new FileInfo(templatePath)))

            {
                Console.WriteLine("Assembling Report Knight Enterprises");
                var ws = pck.Workbook.Worksheets["Data"];
                var a = ws.Cells["A1"].Value;
                int i = 2;
                ws.InsertRow(i, dt.Rows.Count);
                ws.Cells["A1"].LoadFromDataTable(dt, true);
                ws.DeleteRow(dt.Rows.Count + 2);
                pck.Save();
                stop = new TimeSpan(DateTime.Now.Ticks);
                GC.Collect();
                GC.WaitForPendingFinalizers();


            }

        }


        static void assembleKablelink()
        {

            Thread.Sleep(5000);

            TimeSpan stop;
            TimeSpan start = new TimeSpan(DateTime.Now.Ticks);


            DataTable dt = new DataTable();
            using (var con = new SqlConnection(Global.sqlcon))
            using (var cmd = new SqlCommand("ArcherSPKableLink", con))

            using (var da = new SqlDataAdapter(cmd))
            {
                Console.WriteLine("Running Stored Procedure KableLink Communications");
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandTimeout = 600;

                da.Fill(dt);
            }




            string timestamp = DateTime.Now.ToString("MM-dd-yyyy");
            const string templatePath = @"\\tamp20pvfiler09\SHARE1\cfdivix1\Public\Finance\Business Planning\Daily\Reporting Tools\Proj\HotSpot\Template\template.xlsx"; // the path of the template
            string resultPath = @"\\tamp20pvfiler09\SHARE1\cfdivix1\Public\Finance\Business Planning\Daily\Reporting Tools\Proj\HotSpot\Reports\Kablelink\SMB Florida Hotspot Activations Exceptions KableLink" + " " + timestamp + ".xlsx"; // the path of our result
            File.Copy(templatePath, resultPath, true);
            var templateFile = new FileInfo(templatePath);
            var resultsFile = new FileInfo(resultPath);

            if (resultsFile.Exists)
                resultsFile.Delete();



            using (var pck = new ExcelPackage(new FileInfo(resultPath), new FileInfo(templatePath)))

            {
                Console.WriteLine("Assembling Report KableLink Communications");
                var ws = pck.Workbook.Worksheets["Data"];
                var a = ws.Cells["A1"].Value;
                int i = 2;
                ws.InsertRow(i, dt.Rows.Count);
                ws.Cells["A1"].LoadFromDataTable(dt, true);
                ws.DeleteRow(dt.Rows.Count + 2);
                pck.Save();
                stop = new TimeSpan(DateTime.Now.Ticks);
                GC.Collect();
                GC.WaitForPendingFinalizers();


            }

        }

        static void assembleAeon()
        {

            Thread.Sleep(5000);

            TimeSpan stop;
            TimeSpan start = new TimeSpan(DateTime.Now.Ticks);


            DataTable dt = new DataTable();
            using (var con = new SqlConnection(Global.sqlcon))
            using (var cmd = new SqlCommand("ArcherSPAeon", con))

            using (var da = new SqlDataAdapter(cmd))
            {
                Console.WriteLine("Running Stored Procedure AEON INC");
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandTimeout = 600;

                da.Fill(dt);
            }




            string timestamp = DateTime.Now.ToString("MM-dd-yyyy");
            const string templatePath = @"\\tamp20pvfiler09\SHARE1\cfdivix1\Public\Finance\Business Planning\Daily\Reporting Tools\Proj\HotSpot\Template\template.xlsx"; // the path of the template
            string resultPath = @"\\tamp20pvfiler09\SHARE1\cfdivix1\Public\Finance\Business Planning\Daily\Reporting Tools\Proj\HotSpot\Reports\Aeon\SMB Florida Hotspot Activations Exceptions Aeon" + " " + timestamp + ".xlsx"; // the path of our result
            File.Copy(templatePath, resultPath, true);
            var templateFile = new FileInfo(templatePath);
            var resultsFile = new FileInfo(resultPath);

            if (resultsFile.Exists)
                resultsFile.Delete();



            using (var pck = new ExcelPackage(new FileInfo(resultPath), new FileInfo(templatePath)))

            {
                Console.WriteLine("Assembling Report AEON INC");
                var ws = pck.Workbook.Worksheets["Data"];
                var a = ws.Cells["A1"].Value;
                int i = 2;
                ws.InsertRow(i, dt.Rows.Count);
                ws.Cells["A1"].LoadFromDataTable(dt, true);
                ws.DeleteRow(dt.Rows.Count + 2);
                pck.Save();
                stop = new TimeSpan(DateTime.Now.Ticks);
                GC.Collect();
                GC.WaitForPendingFinalizers();


            }

        }


        static void assembleJaguar()
        {

            Thread.Sleep(5000);

            TimeSpan stop;
            TimeSpan start = new TimeSpan(DateTime.Now.Ticks);


            DataTable dt = new DataTable();
            using (var con = new SqlConnection(Global.sqlcon))
            using (var cmd = new SqlCommand("ArcherSPJaguar", con))

            using (var da = new SqlDataAdapter(cmd))
            {
                Console.WriteLine("Running Stored Procedure JAGUAR TECHNOLOGIES INC");
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandTimeout = 600;

                da.Fill(dt);
            }




            string timestamp = DateTime.Now.ToString("MM-dd-yyyy");
            const string templatePath = @"\\tamp20pvfiler09\SHARE1\cfdivix1\Public\Finance\Business Planning\Daily\Reporting Tools\Proj\HotSpot\Template\template.xlsx"; // the path of the template
            string resultPath = @"\\tamp20pvfiler09\SHARE1\cfdivix1\Public\Finance\Business Planning\Daily\Reporting Tools\Proj\HotSpot\Reports\Jaguar\SMB Florida Hotspot Activations Exceptions Jaguar" + " " + timestamp + ".xlsx"; // the path of our result
            File.Copy(templatePath, resultPath, true);
            var templateFile = new FileInfo(templatePath);
            var resultsFile = new FileInfo(resultPath);

            if (resultsFile.Exists)
                resultsFile.Delete();



            using (var pck = new ExcelPackage(new FileInfo(resultPath), new FileInfo(templatePath)))

            {
                Console.WriteLine("Assembling Report JAGUAR TECHNOLOGIES INC");
                var ws = pck.Workbook.Worksheets["Data"];
                var a = ws.Cells["A1"].Value;
                int i = 2;
                ws.InsertRow(i, dt.Rows.Count);
                ws.Cells["A1"].LoadFromDataTable(dt, true);
                ws.DeleteRow(dt.Rows.Count + 2);
                pck.Save();
                stop = new TimeSpan(DateTime.Now.Ticks);
                GC.Collect();
                GC.WaitForPendingFinalizers();


            }

        }






        static void MailFileInhouse()
        {
            Thread.Sleep(8000);
            try
            {
                Console.WriteLine("Emailing Inhouse Managers.....");
                //Console.WriteLine(ToValue + ',' + ToCC);
                //Console.ReadKey();
                MailMessage mailMessage = new MailMessage();
                mailMessage.From = new MailAddress("DL-FL-Region-Operations-Analytics@charter.com");
                //mailMessage.To.Add("Eric.Fiallo@charter.com");
                mailMessage.To.Add("Dale.Peterson@charter.com,Andrew.Cmar@charter.com,Robert.Carter@charter.com,Doug.Delk@charter.com");
                mailMessage.CC.Add("Allen.K.Smith@charter.com,Nicole.Buchanan@charter.com,Eric.Fiallo@charter.com");
                SmtpClient smtpClient = new SmtpClient();
                mailMessage.Subject = "Spectrum Hotspots: Florida Region Successful Activations and Exceptions for Last 7 Days";
                //mailMessage.Body = "Spectrum Hotspots: Florida Region Successful Activations and Exceptions for Last 7 Days Aeon";




                System.Net.Mail.Attachment attachment;
                var directory = new DirectoryInfo(@"\\tamp20pvfiler09\SHARE1\cfdivix1\Public\Finance\Business Planning\Daily\Reporting Tools\Proj\HotSpot\Reports");
                var att = (from f in directory.GetFiles()
                           orderby f.LastWriteTime descending
                           select f).First();
                attachment = new System.Net.Mail.Attachment(att.FullName);
                mailMessage.Attachments.Add(attachment);




                smtpClient = new SmtpClient("Mailrelay.chartercom.com");
                smtpClient.UseDefaultCredentials = false;
                smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtpClient.DeliveryFormat = SmtpDeliveryFormat.SevenBit;
                smtpClient.Send(mailMessage);
            }
            catch (Exception ex)
            {
                string path = @"C:\Users\P2162735\ErrorLog\Error.txt";  // file path
                using (StreamWriter sw = new StreamWriter(path, true))
                { // If file exists, text will be appended ; otherwise a new file will be created
                    sw.Write(string.Format("Message: {0}<br />{1}StackTrace :{2}{1}Date :{3}{1}-----------------------------------------------------------------------------{1}", ex.Message, Environment.NewLine, ex.StackTrace, DateTime.Now.ToString()));
                }

            }






        }


        static void MailFileAeon()
        {
            Thread.Sleep(5000);
            try
            {
                Console.WriteLine("Emailing AEON Managers.....");
                //Console.WriteLine(ToValue + ',' + ToCC);
                //Console.ReadKey();
                MailMessage mailMessage = new MailMessage();
                mailMessage.From = new MailAddress("DL-FL-Region-Operations-Analytics@charter.com");
                //mailMessage.To.Add("Eric.Fiallo@charter.com");
                mailMessage.To.Add("rkohn@aeon.tech,CLacore@aeon.tech,arosario@aeon.tech,MDierking@aeon.tech");
                mailMessage.CC.Add("Allen.K.Smith@charter.com,Nicole.Buchanan@charter.com,Eric.Fiallo@charter.com");
                SmtpClient smtpClient = new SmtpClient();
                mailMessage.Subject = "Spectrum Hotspots: Florida Region Successful Activations and Exceptions Aeon";
                //mailMessage.Body = "Spectrum Hotspots: Florida Region Successful Activations and Exceptions Aeon";




                System.Net.Mail.Attachment attachment;
                var directory = new DirectoryInfo(@"\\tamp20pvfiler09\SHARE1\cfdivix1\Public\Finance\Business Planning\Daily\Reporting Tools\Proj\HotSpot\Reports\Aeon");
                var att = (from f in directory.GetFiles()
                           orderby f.LastWriteTime descending
                           select f).First();
                attachment = new System.Net.Mail.Attachment(att.FullName);
                mailMessage.Attachments.Add(attachment);




                smtpClient = new SmtpClient("Mailrelay.chartercom.com");
                smtpClient.UseDefaultCredentials = false;
                smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtpClient.DeliveryFormat = SmtpDeliveryFormat.SevenBit;
                smtpClient.Send(mailMessage);
            }
            catch (Exception ex)
            {
                string path = @"C:\Users\P2162735\ErrorLog\Error.txt";  // file path
                using (StreamWriter sw = new StreamWriter(path, true))
                { // If file exists, text will be appended ; otherwise a new file will be created
                    sw.Write(string.Format("Message: {0}<br />{1}StackTrace :{2}{1}Date :{3}{1}-----------------------------------------------------------------------------{1}", ex.Message, Environment.NewLine, ex.StackTrace, DateTime.Now.ToString()));
                }

            }






        }




        static void MailFileKablelink()
        {
            Thread.Sleep(8000);
            try
            {
                Console.WriteLine("Emailing Kablelink Managers.....");
                //Console.WriteLine(ToValue + ',' + ToCC);
                //Console.ReadKey();
                MailMessage mailMessage = new MailMessage();
                mailMessage.From = new MailAddress("DL-FL-Region-Operations-Analytics@charter.com");
                mailMessage.To.Add("dl-managers@kablelink.com,rwalker@kablelink.com");
                mailMessage.CC.Add("Allen.K.Smith@charter.com,Nicole.Buchanan@charter.com,Eric.Fiallo@charter.com");
                SmtpClient smtpClient = new SmtpClient();
                mailMessage.Subject = "Spectrum Hotspots: Florida Region Successful Activations and Exceptions Kablelink";
                //mailMessage.Body = "Spectrum Hotspots: Florida Region Successful Activations and Exceptions Kablelink";




                System.Net.Mail.Attachment attachment;
                var directory = new DirectoryInfo(@"\\tamp20pvfiler09\SHARE1\cfdivix1\Public\Finance\Business Planning\Daily\Reporting Tools\Proj\HotSpot\Reports\Kablelink");
                var att = (from f in directory.GetFiles()
                           orderby f.LastWriteTime descending
                           select f).First();
                attachment = new System.Net.Mail.Attachment(att.FullName);
                mailMessage.Attachments.Add(attachment);




                smtpClient = new SmtpClient("Mailrelay.chartercom.com");
                smtpClient.UseDefaultCredentials = false;
                smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtpClient.DeliveryFormat = SmtpDeliveryFormat.SevenBit;
                smtpClient.Send(mailMessage);
            }
            catch (Exception ex)
            {
                string path = @"C:\Users\P2162735\ErrorLog\Error.txt";  // file path
                using (StreamWriter sw = new StreamWriter(path, true))
                { // If file exists, text will be appended ; otherwise a new file will be created
                    sw.Write(string.Format("Message: {0}<br />{1}StackTrace :{2}{1}Date :{3}{1}-----------------------------------------------------------------------------{1}", ex.Message, Environment.NewLine, ex.StackTrace, DateTime.Now.ToString()));
                }

            }






        }




<<<<<<< HEAD
        static void MailFileJag()
=======
         static void MailFileJag()
>>>>>>> 5381eb514bdf63a3870fe0d57232c66f9dca89a5
        {
            Thread.Sleep(8000);
            try
            {
                Console.WriteLine("Emailing Jag Managers.....");
                //Console.WriteLine(ToValue + ',' + ToCC);
                //Console.ReadKey();
                MailMessage mailMessage = new MailMessage();
                mailMessage.From = new MailAddress("DL-FL-Region-Operations-Analytics@charter.com");
                mailMessage.To.Add("jagnorthmgmt@jaguartechnologies.com");
                mailMessage.CC.Add("Allen.K.Smith@charter.com,Nicole.Buchanan@charter.com,Eric.Fiallo@charter.com");
                SmtpClient smtpClient = new SmtpClient();
                mailMessage.Subject = "Spectrum Hotspots: Florida Region Successful Activations and Exceptions Jaguar";
                //mailMessage.Body = "Spectrum Hotspots: Florida Region Successful Activations and Exceptions Kablelink";




                System.Net.Mail.Attachment attachment;
                var directory = new DirectoryInfo(@"\\tamp20pvfiler09\SHARE1\cfdivix1\Public\Finance\Business Planning\Daily\Reporting Tools\Proj\HotSpot\Reports\Jaguar");
                var att = (from f in directory.GetFiles()
                           orderby f.LastWriteTime descending
                           select f).First();
                attachment = new System.Net.Mail.Attachment(att.FullName);
                mailMessage.Attachments.Add(attachment);


<<<<<<< HEAD


                smtpClient = new SmtpClient("Mailrelay.chartercom.com");
                smtpClient.UseDefaultCredentials = false;
                smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtpClient.DeliveryFormat = SmtpDeliveryFormat.SevenBit;
                smtpClient.Send(mailMessage);
            }
            catch (Exception ex)
            {
                string path = @"C:\Users\P2162735\ErrorLog\Error.txt";  // file path
                using (StreamWriter sw = new StreamWriter(path, true))
                { // If file exists, text will be appended ; otherwise a new file will be created
                    sw.Write(string.Format("Message: {0}<br />{1}StackTrace :{2}{1}Date :{3}{1}-----------------------------------------------------------------------------{1}", ex.Message, Environment.NewLine, ex.StackTrace, DateTime.Now.ToString()));
                }

            }

=======
>>>>>>> 5381eb514bdf63a3870fe0d57232c66f9dca89a5


                smtpClient = new SmtpClient("Mailrelay.chartercom.com");
                smtpClient.UseDefaultCredentials = false;
                smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtpClient.DeliveryFormat = SmtpDeliveryFormat.SevenBit;
                smtpClient.Send(mailMessage);
            }
            catch (Exception ex)
            {
                string path = @"C:\Users\P2162735\ErrorLog\Error.txt";  // file path
                using (StreamWriter sw = new StreamWriter(path, true))
                { // If file exists, text will be appended ; otherwise a new file will be created
                    sw.Write(string.Format("Message: {0}<br />{1}StackTrace :{2}{1}Date :{3}{1}-----------------------------------------------------------------------------{1}", ex.Message, Environment.NewLine, ex.StackTrace, DateTime.Now.ToString()));
                }

            }






        }




        }
        public static bool IsFileinUse(string path)
        {
            FileStream streams = null;

            try
            {
                Stream stream = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            }
            catch (IOException)
            {
                //the file is unavailable because it is:
                //still being written to
                //or being processed by another thread
                //or does not exist (has already been processed)
                return true;
            }
            finally
            {
                if (streams != null)
                    streams.Close();
            }
            return false;
        }






    }















}














