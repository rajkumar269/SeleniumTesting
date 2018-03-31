using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using NUnit.Framework;
using System.Net.Mail;


namespace Setup
{
    [SetUpFixture]
    public class SetupClass
    {
        public static List<TestCaseSettingList> tl = new List<TestCaseSettingList>();

        public static DataSet ds = new DataSet();
        public static DataSet Outputds = new DataSet();
        public static string masterPath;
        public static string BrowserName;       
        public static string TestCaseDocumentPath = AppDomain.CurrentDomain.BaseDirectory + @"TestCase\TestCaseInput.xls";
        public static string ResultTestCaseDocumentPath = AppDomain.CurrentDomain.BaseDirectory + @"TestCase\TestingReport.xls";
        
        public static string sScreenShotPath = AppDomain.CurrentDomain.BaseDirectory + @"ScreenShots\";
        
        public static string TestModules;
        public static string TestPerson;
        public static DataSet OutputDS
        {
            get
            {
                return Outputds;
            }
            set
            {
                Outputds = value;
            }
        }
        public static DataSet DS
        {
            get
            {
                return ds;
            }
            set
            {
                ds = value;
            }
        }

        public string MasterPath
        {
            get
            {
                return masterPath;
            }
            set
            {
                masterPath = value;
            }
        }

       
        public List<TestCaseSettingList> GetList()
        {
            return tl;
        }  



        [SetUp]
        public void ReadingValuesFromExcel()
        {
            try
            {
                
                //Added by David to delete the screen shot while running the 
                Console.WriteLine("Screenshots directory cleared.....");
                foreach (string filePath in Directory.GetFiles(sScreenShotPath))
                    File.Delete(filePath);

                DS = new DataSet();
                               
                var connectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0; data source={0}; Extended Properties=Excel 8.0;", TestCaseDocumentPath);
                OleDbConnection con = new OleDbConnection(connectionString);               
                con.Open();
                Console.WriteLine("Excel DataReader Connection Open..");
                OleDbCommand cmd = new OleDbCommand();
                OleDbDataAdapter oleda = new OleDbDataAdapter();

                DataTable dt = new DataTable();
                dt = con.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, null);
                int rowsCount = dt.Rows.Count;
                Console.WriteLine("Reading Values from the Excel..."); 
                for (int i = 0; i < rowsCount; i++)
                {
                    string workSheetName = (string)dt.Rows[i]["TABLE_NAME"];
                    //Console.WriteLine(workSheetName);
                    cmd.Connection = con;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "SELECT * FROM [" + workSheetName + "]";

                    oleda = new OleDbDataAdapter(cmd);

                    if (workSheetName.Contains("Sheet1") || TestModules.Contains(workSheetName.Replace("$","")))
                    {
                        oleda.Fill(ds, workSheetName);
                    }                    

                    if (workSheetName.Contains("Sheet1"))
                    {
                        SetProperties(ds, workSheetName);
                    }
                    
                }
                con.Close();
                Outputds = ds.Copy();
                foreach (DataTable dts in Outputds.Tables)
                {
                    if (dts.TableName.Contains("Sheet1"))
                    {
                        continue;
                    }
                    else
                    {
                        dts.Columns.Add("Developer Status");
                        dts.Columns.Add("TestCase Status");      
                        dts.Columns.Add("Internet Explorer");
                        dts.Columns.Add("Firefox");                        
                        dts.Columns.Add("Chrome");                                       
                                          
                        //dts.Columns.Add();
                        //dts.Columns.Add();
                    }
                }
                Console.WriteLine("Excel Data Reader Connection closed...");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message.ToString());
            }
        }


        public void SetProperties(DataSet dts, string TableName)
        {
            DataTable Properties = dts.Tables[TableName];            
            masterPath = Properties.Rows[0][0].ToString();
            TestModules = Properties.Rows[1][0].ToString();
            TestPerson = Properties.Rows[2][0].ToString();
            
        }
        [TearDown]
        public void WriteResultsToExcel()
        {
            
            StreamWriter fs = new StreamWriter(ResultTestCaseDocumentPath, false);
            
            fs.WriteLine("<?xml version=\"1.0\"?>");
            fs.WriteLine("<?mso-application progid=\"Excel.Sheet\"?>");
            fs.WriteLine("<ss:Workbook xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\">");
            fs.WriteLine("    <ss:Styles>");
            fs.WriteLine("        <ss:Style ss:ID=\"1\">");
            fs.WriteLine("           <ss:Font ss:Bold=\"1\" />");            
            fs.WriteLine("        </ss:Style>");

            fs.WriteLine("    </ss:Styles>");
            foreach (DataTable OutputTable in Outputds.Tables)
            {
                fs.WriteLine("    <ss:Worksheet ss:Name=\"" + OutputTable.TableName + "\">");
                fs.WriteLine("        <ss:Table>");
                for (int x = 0; x <= OutputTable.Columns.Count - 1; x++)
                {
                    fs.WriteLine("            <ss:Column ss:Width=\"{0}\"/>", 100);

                }
                fs.WriteLine("            <ss:Row ss:StyleID=\"1\">");
                foreach (DataColumn column in OutputTable.Columns)
                {
                    fs.WriteLine("                <ss:Cell>");
                    fs.WriteLine(string.Format("                   <ss:Data ss:Type=\"String\">{0}</ss:Data>", column.ColumnName));
                    fs.WriteLine("                </ss:Cell>");
                }
                fs.WriteLine("            </ss:Row>");
                for (int intRow = 0; intRow <= OutputTable.Rows.Count - 1; intRow++)
                {
                    
                    fs.WriteLine(string.Format("            <ss:Row ss:Height =\"{0}\">", 15));
                   
                    for (int intCol = 0; intCol <= OutputTable.Columns.Count - 1; intCol++)
                    {
                        fs.WriteLine("                <ss:Cell>");
                        fs.WriteLine(string.Format("                   <ss:Data ss:Type=\"String\">{0}</ss:Data>", OutputTable.Rows[intRow][intCol].ToString()));
                        fs.WriteLine("                </ss:Cell>");
                    }
                    fs.WriteLine("            </ss:Row>");
                }
                fs.WriteLine("        </ss:Table>");
                fs.WriteLine("    </ss:Worksheet>");
            }

            fs.WriteLine("</ss:Workbook>");
            fs.Close();


            string LogonUserName = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            Send("The selenium test project was built and run. Kindly check the report attached.", "Selenium Test Excel Report", "davidrajkumar.j@dsrc.co.in");
        }
        private static void Send(string emailbody, string emailSub, string ToAddress)
        {
            MailMessage mail = new MailMessage();
            SmtpClient client = new SmtpClient();
            client.Port = 2525;
            client.Host = "192.168.4.101";
            client.Credentials = new System.Net.NetworkCredential("prasanthk", "dsrc12345");
            mail.From = new MailAddress("davidrajkumar.j@dsrc.co.in");
            mail.To.Add(ToAddress);
            mail.Body = emailbody;
            mail.Subject = emailSub;
            System.Net.Mail.Attachment attachment;
            attachment = new System.Net.Mail.Attachment(@"C:\Users\davidrajkumar.j\.jenkins\workspace\GenericPortalSeleniumTesting\SeleniumTestingProject\bin\Debug\TestCase\TestingReport.xls");
            mail.Attachments.Add(attachment);
            mail.IsBodyHtml = true;
            client.Send(mail);
        }
    }
}