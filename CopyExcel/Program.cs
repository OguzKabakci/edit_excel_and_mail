using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace CopyExcel
{
    class Program
    {
        // Read data
        static int[] read_data = new int[41];

        static void Main(string[] args)
        {
            ReadExcel();
            //WriteExcel();
            //SendMail();
        }
        public static void ReadExcel()
        {
            //Get month number
            int monthNumber = DateTime.Now.Month;

            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Users\vjkyky\Documents\Volkan\Visual_excel_copy\SAYAÇ 2019.xls");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[monthNumber];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            //Read data
            for (int i = 0; i < 41; i++)
            {
                read_data[i] = (int)xlRange.Cells[i + 4, 11].Value2;
            }

            //Close workbook
            xlWorkbook.Close();
        }
        public static void WriteExcel()
        {
            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Users\vjkyky\Documents\Volkan\Visual_excel_copy\TUKETIM_ORIGINAL_12_19_tek_zamanlı_tarife.xls");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[2];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            //Write data
            for (int i = 0; i < 41; i++)
            {
                xlWorksheet.Cells[i + 2, "H"] = read_data[i];
            }
            xlWorkbook.RefreshAll();
            // Save and close workbook
            xlWorkbook.Save();
            xlWorkbook.Close();
            
        }

        public static void SendMail()
        {
            var ol = new Outlook.Application();
            Outlook.MailItem mail = ol.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
            mail.To = "oguz.kabakci94@hotmail.com; VolkanDalkilic@windowslive.com";
            mail.Subject = "Enter subject here";
            mail.Body = "Enter body here";
            mail.Attachments.Add(@"C:\Users\vjkyky\Documents\Volkan\Visual_excel_copy\TUKETIM_ORIGINAL_12_19_tek_zamanlı_tarife.xls");
            mail.Send();
            mail = null;
            ol = null;
        }
    }


}
