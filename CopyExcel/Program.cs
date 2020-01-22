using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace CopyExcel
{
    class Program
    {
        // Read data
        static int[] read_data = new int[41];
        static int dataNumber;
        static string writeExcelPath;

        static void Main(string[] args)
        {
            //Get data number to be read
            dataNumber = int.Parse(ConfigurationManager.AppSettings["dataNumber"]);
            //Get write excel path
            writeExcelPath = ConfigurationManager.AppSettings["writeExcelPath"];
            //Read excel data
            ReadExcel();
            Console.WriteLine("Data is read.");
            //Write data to excel
            WriteExcel();
            Console.WriteLine("Data is written and excel is refreshed.");
            //Send mail
            SendMail();
            Console.WriteLine("Mail is sent.");
            Console.WriteLine("Pres any key to close the program...");
            Console.ReadKey();
            Environment.Exit(0);
        }
        public static void ReadExcel()
        {
            //Get month number
            int monthNumber = DateTime.Now.Month;

            //get read excel data
            string readExcelPath = ConfigurationManager.AppSettings["readExcelPath"];
            int readExcelStartRow = int.Parse(ConfigurationManager.AppSettings["readExcelStartRow"]);
            int readExcelStartLine = int.Parse(ConfigurationManager.AppSettings["readExcelStartLine"]);

            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(readExcelPath);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[monthNumber];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            //Read data
            for (int i = 0; i < dataNumber; i++)
            {
                read_data[i] = (int)xlRange.Cells[i + readExcelStartLine, readExcelStartRow].Value2;
            }

            //Close workbook
            xlWorkbook.Close();
        }
        public static void WriteExcel()
        {
            //Get write excel data
            int writeExcelStartRow = int.Parse(ConfigurationManager.AppSettings["writeExcelStartRow"]);
            int writeExcelStartLine = int.Parse(ConfigurationManager.AppSettings["writeExcelStartLine"]);
            int writeExcelSheetNumber = int.Parse(ConfigurationManager.AppSettings["writeExcelSheetNumber"]);

            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(writeExcelPath);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[writeExcelSheetNumber];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            //Write data
            for (int i = 0; i < dataNumber; i++)
            {
                xlWorksheet.Cells[i + writeExcelStartLine, writeExcelStartRow] = read_data[i];
            }

            //Refresh the workbook to run macros
            xlWorkbook.RefreshAll();

            // Save and close workbook
            xlWorkbook.Save();
            xlWorkbook.Close();
            
        }

        public static void SendMail()
        {
            //Get mail data
            string mailReceivers = ConfigurationManager.AppSettings["mailReceivers"];
            string mailSubject = ConfigurationManager.AppSettings["mailSubject"];
            string mailBody = ConfigurationManager.AppSettings["mailBody"];

            //Prepare the mail
            var ol = new Outlook.Application();
            Outlook.MailItem mail = ol.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
            mail.To = mailReceivers;
            mail.Subject = mailSubject;
            mail.Body = mailBody;
            mail.Attachments.Add(writeExcelPath);
            mail.Send();

            //Clear
            mail = null;
            ol = null;
        }
    }


}
