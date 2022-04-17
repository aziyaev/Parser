
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;



namespace Parser
{
    public class XlFile : IXlFile
    {
        public static List<Note> Sheet { get; set; } = new List<Note>();
        private Excel.Application WorkExcel { get; set; }

        public XlFile()
        {
            WorkExcel = new Excel.Application();
            OpenTable();
        }

        public XlParser LoadTable(out string message)
        {
            WorkExcel = new Excel.Application();
            try
            {
                message = "Успешно";
                return new XlParser(WorkExcel);
            }
            catch (Exception)
            {
                message = "Ошибка";
                return null;
            }
        }

        public void OpenTable()
        {
            string filepath = Path.GetFullPath("..").Substring(0, Path.GetFullPath("..").Length - 3) + "listAlert.xlsx";
            

            if (!File.Exists(filepath))
            {
                if (!CreateTable(filepath))
                {
                    //idk
                    return;
                }
            }

            

            Excel.Workbook workBook = WorkExcel.Workbooks.Open(filepath);
            Excel.Worksheet workSheet = (Excel.Worksheet)workBook.Sheets[1];

            List<Note> notes = ConvertTable(workSheet);

            foreach(Note note in notes)
            {
                Sheet.Add(note);
            }

            if (workSheet != null) Marshal.ReleaseComObject(workSheet);
            if (workBook != null)
            {
                workBook.Close(false);
                Marshal.ReleaseComObject(workBook);
                workBook = null;
            }
            if(WorkExcel != null)
            {
                WorkExcel.Quit();
                Marshal.ReleaseComObject(WorkExcel);
                WorkExcel = null;
            }
        }

        public void SaveTable()
        {
            WorkExcel = new Excel.Application();
            string filepath = Path.GetFullPath("..").Substring(0, Path.GetFullPath("..").Length - 3) + "listAlert.xlsx";

            Excel.Workbook workBook = WorkExcel.Workbooks.Add();
            Excel.Worksheet workSheet = (Excel.Worksheet)workBook.ActiveSheet;
            int headerIndex = 1;

            Excel.Range headerTextRange = workSheet.get_Range("A1", "E1");
            headerTextRange.WrapText = true;

            headerTextRange = workSheet.get_Range("F1", "H1");
            headerTextRange.WrapText = true;

            headerTextRange = workSheet.get_Range("I1", "J1");
            headerTextRange.WrapText = true;

            workSheet.Cells[headerIndex, 1] = "Общая информация";
            workSheet.Cells[headerIndex, 6] = "Последствия";
            workSheet.Cells[headerIndex, 9] = "Дополнительно";


            headerIndex++;
            workSheet.Cells[headerIndex, 1] = "Идентификатор УБИ";
            workSheet.Cells[headerIndex, 2] = "Наименование УБИ";
            workSheet.Cells[headerIndex, 3] = "Описание";
            workSheet.Cells[headerIndex, 4] = "Источник угрозы (характеристика и потенциал нарушителя)";
            workSheet.Cells[headerIndex, 5] = "Объект воздействия";
            workSheet.Cells[headerIndex, 6] = "Нарушение конфиденциальности";
            workSheet.Cells[headerIndex, 7] = "Нарушение целостности";
            workSheet.Cells[headerIndex, 8] = "Нарушение доступности";
            workSheet.Cells[headerIndex, 9] = "Дата включения угрозы в БнД УБИ";
            workSheet.Cells[headerIndex, 10] = "Дата последнего изменения данных";

            int rowIndex = 3;
            foreach(Note note in Sheet)
            {
                workSheet.Cells[rowIndex, 1] = note.Id;
                workSheet.Cells[rowIndex, 2] = note.Name;
                workSheet.Cells[rowIndex, 3] = note.Description;
                workSheet.Cells[rowIndex, 4] = note.Source;
                workSheet.Cells[rowIndex, 5] = note.Threat;
                workSheet.Cells[rowIndex, 6] = note.IsNotConfidential;
                workSheet.Cells[rowIndex, 7] = note.IsComplete;
                workSheet.Cells[rowIndex, 8] = note.IsAccessible;
                workSheet.Cells[rowIndex, 9] = note.DateIn;
                workSheet.Cells[rowIndex, 10] = note.DateRewrite;
                rowIndex++;
            }

            using (SaveFileDialog exportFile = new SaveFileDialog())
            {
                exportFile.Title = "Экспорт файла";
                exportFile.Filter = "Microsoft Office Excel Workbook(*.xlsx)|*.xlsx";

                if(DialogResult.OK == exportFile.ShowDialog())
                {
                    filepath = exportFile.FileName;
                    workBook.SaveAs(filepath, Excel.XlFileFormat.xlOpenXMLWorkbook, 
                        Missing.Value, Missing.Value, false, false, 
                        Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlUserResolution, 
                        true, Missing.Value, Missing.Value, Missing.Value);
                    workBook.Saved = true;


                }
            }
        }

        public string UpdateTable()
        {
            
            var table = LoadTable(out string message);

            if (table == null)
            {
                return message;
            }

            
            Excel.Worksheet workSheet = (Excel.Worksheet)table.WorkBook.Sheets[1];
            List<Note> notes = ConvertTable(workSheet);
            List<Note> notesOld = new List<Note>();
            List<Note> notesNew = new List<Note>();

            foreach (Note noteSheet in Sheet)
            {
                foreach(Note noteTemp in notes)
                {
                    if(noteSheet.Id == noteTemp.Id && (noteSheet.Name != noteTemp.Name 
                        || noteSheet.Description != noteTemp.Description 
                        || noteSheet.Source != noteTemp.Source 
                        || noteSheet.Threat != noteTemp.Threat 
                        || noteSheet.IsNotConfidential != noteTemp.IsNotConfidential 
                        || noteSheet.IsComplete != noteTemp.IsComplete
                        || noteSheet.IsAccessible != noteTemp.IsAccessible))
                    {
                        ChangeInfoWindow.notesOld.Add(new Note(noteSheet.Id, noteSheet.IdInfo, 
                            noteSheet.Name, noteSheet.Description, noteSheet.Source, 
                            noteSheet.Threat, noteSheet.IsNotConfidential, noteSheet.IsComplete, 
                            noteSheet.IsAccessible, noteSheet.DateIn, noteSheet.DateRewrite));
                        ChangeInfoWindow.notesNew.Add(noteTemp);
                        noteSheet.Name = noteTemp.Name;
                        noteSheet.Description = noteTemp.Description;
                        noteSheet.Source = noteTemp.Source;
                        noteSheet.Threat = noteTemp.Threat;
                        noteSheet.IsNotConfidential = noteTemp.IsNotConfidential;
                        noteSheet.IsComplete = noteTemp.IsComplete;
                        noteSheet.IsAccessible = noteTemp.IsAccessible;
                        noteSheet.DateIn = noteTemp.DateIn;
                        noteSheet.DateRewrite = noteTemp.DateRewrite;
                        break;
                    }
                }
            }
            if (workSheet != null) Marshal.ReleaseComObject(workSheet);
            if (table.WorkBook != null)
            {
                table.WorkBook.Close(false);
                Marshal.ReleaseComObject(table.WorkBook);
                table.WorkBook = null;
            }
            if (WorkExcel != null)
            {
                WorkExcel.Quit();
                Marshal.ReleaseComObject(WorkExcel);
                WorkExcel = null;
            }


            return message;
        }

        private bool CreateTable(string filepath)
        {
            var table = LoadTable(out string message);

            if (table != null)
            {
                table.WorkBook.SaveAs(filepath);
                return true;
            }

            return false;
            //Create empty table

        }

        private List<Note> ConvertTable(Excel.Worksheet workSheet)
        {
            List<Note> list = new List<Note>();

            var lastCell = workSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int lastRow = lastCell.Row;

            for (int i = 3; i < lastRow; i++)
            {
                try
                {
                    int id = Convert.ToInt32(workSheet.Cells[i, 1].Text.ToString());
                    string idInfo = $"УБИ.{id.ToString()}";
                    string name = workSheet.Cells[i, 2].Text.ToString();
                    string description = workSheet.Cells[i, 3].Text.ToString();
                    string source = workSheet.Cells[i, 4].Text.ToString();
                    string threat = workSheet.Cells[i, 5].Text.ToString();
                    string isNotConfidential = "Нет";
                    if (Convert.ToInt32(workSheet.Cells[i, 6].Text.ToString()) == 1)
                    {
                        isNotConfidential = "Да";
                    }
                    string isComplete = "Нет";
                    if (Convert.ToInt32(workSheet.Cells[i, 7].Text.ToString()) == 1)
                    {
                        isComplete = "Да";
                    }
                    string isAccessible = "Нет";
                    if (Convert.ToInt32(workSheet.Cells[i, 8].Text.ToString()) == 1)
                    {
                        isAccessible = "Да";
                    }
                    string dateIn = workSheet.Cells[i, 9].Text.ToString();
                    string dateRewrite = workSheet.Cells[i, 9].Text.ToString();
                    list.Add(new Note(id, idInfo, name, description, source, threat, isNotConfidential, isComplete, isAccessible, dateIn, dateRewrite));
                }
                catch (Exception)
                {
                    continue;
                }
            }
            return list;
        }
    }
}
