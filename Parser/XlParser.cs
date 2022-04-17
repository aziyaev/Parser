using ClosedXML.Excel;
using Grpc.Core;
using System;
using System.IO;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;
using System.Collections.Generic;
using System.Reflection;

namespace Parser
{
    public class XlParser
    {
        public Excel.Application WorkExcel { get; set; }
        public XLWorkbook Workbook { get; set; }


        public XlParser(bool isDownload = false)
        {
            string link = "https://bdu.fstec.ru/files/documents/thrlist.xlsx";
            string path = Path.GetFullPath(".") + "\\";

            if (!File.Exists(path + "listAlert.xlsx") || isDownload)
            {
                WorkExcel = new Excel.Application();
                Excel.Workbook wb = WorkExcel.Workbooks.Open(link);
                if (isDownload) 
                {
                    wb.SaveAs(path + "newListAlert.xlsx", Excel.XlFileFormat.xlOpenXMLWorkbook,
                        Missing.Value, Missing.Value, false, false,
                        Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlUserResolution,
                        true, Missing.Value, Missing.Value, Missing.Value);
                    wb.SaveAs(path + "listAlert.xlsx", Excel.XlFileFormat.xlOpenXMLWorkbook,
                        Missing.Value, Missing.Value, false, false,
                        Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlUserResolution,
                        true, Missing.Value, Missing.Value, Missing.Value);
                }  
                else wb.SaveAs(path + "listAlert.xlsx", Excel.XlFileFormat.xlOpenXMLWorkbook,
                        Missing.Value, Missing.Value, false, false,
                        Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlUserResolution,
                        true, Missing.Value, Missing.Value, Missing.Value);
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
        }
    }
}
