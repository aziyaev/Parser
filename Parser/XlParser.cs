using ClosedXML.Excel;
using Grpc.Core;
using System;
using System.IO;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;
using System.Collections.Generic;


namespace Parser
{
    public class XlParser
    {
        public Excel.Application WorkExcel { get; set; }
        public XLWorkbook Workbook { get; set; }

        public XlParser(bool isDownload = false)
        {
            string link = "https://bdu.fstec.ru/files/documents/thrlist.xlsx";
            string path = Path.GetFullPath("..").Substring(0, Path.GetFullPath("..").Length - 3);

            if (!File.Exists(path + "listAlert.xlsx") || isDownload)
            {
                WorkExcel = new Excel.Application();
                Excel.Workbook wb = WorkExcel.Workbooks.Open(link);
                if (isDownload) 
                {
                    wb.SaveAs(path + "newListAlert.xlsx");
                    wb.SaveAs(path + "listAlert.xlsx");
                }  
                else wb.SaveAs(path + "listAlert.xlsx");
                wb.Close();
                WorkExcel.Quit();
            }

            if (isDownload)
            {
                Workbook = new XLWorkbook(path + "newListAlert.xlsx");
            }
            else
            {
                Workbook = new XLWorkbook(path + "listAlert.xlsx");
            }
            //OpenLink(objWorkExcel, link, out Excel.Workbook workBook);
            //WorkBook = workExcel.Workbooks.Open(link);
        }
    }
}
