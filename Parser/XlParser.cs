using System;
using System.IO;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace Parser
{
    public class XlParser
    {
        public Excel.Workbook WorkBook { get; set; }

        public XlParser(Excel.Application objWorkExcel)
        {
            string link = "https://bdu.fstec.ru/files/documents/thrlist.xlsx";
            OpenLink(objWorkExcel, link, out Excel.Workbook workBook);
            WorkBook = workBook;
        }

        public void OpenLink(Excel.Application workExcel, string link, out Excel.Workbook workBook)
        {
            workBook = workExcel.Workbooks.Open(link);
        }

    }
}
