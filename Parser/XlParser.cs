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

        public XlParser()
        {
            string link = "https://bdu.fstec.ru/files/documents/thrlist.xlsx";
            string path = Path.GetFullPath("..").Substring(0, Path.GetFullPath("..").Length - 3) + "listAlert.xlsx";

            if (!File.Exists(path))
            {
                WorkExcel = new Excel.Application();
                Excel.Workbook wb = WorkExcel.Workbooks.Open(link);
                wb.SaveAs(path);
                wb.Close();
                WorkExcel.Quit();
            }

            Workbook = new XLWorkbook(path);
            //OpenLink(objWorkExcel, link, out Excel.Workbook workBook);
            //WorkBook = workExcel.Workbooks.Open(link);
        }

        public XlParser(bool isDowload)
        {
            string link = "https://bdu.fstec.ru/files/documents/thrlist.xlsx";
            string path = Path.GetFullPath("..").Substring(0, Path.GetFullPath("..").Length - 3) + "listAlert.xlsx";

            if (isDowload)
            {
                WorkExcel = new Excel.Application();
                Excel.Workbook wb = WorkExcel.Workbooks.Open(link);
                wb.SaveAs(path);
                wb.Close();
                WorkExcel.Quit();
            }

            Workbook = new XLWorkbook(path);
            //OpenLink(objWorkExcel, link, out Excel.Workbook workBook);
            //WorkBook = workExcel.Workbooks.Open(link);
        }
    }
}
