
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using ClosedXML.Excel;



namespace Parser
{
    public class XlFile : IXlFile
    {
        public static List<Note> Sheet { get; set; } = new List<Note>();

        public XlFile()
        {
            //WorkExcel = new Excel.Application();
            OpenTable();
        }

        public XlParser LoadTable(out string message, bool isDowload = false)
        {
            //WorkExcel = new Excel.Application();
            if (isDowload)
            {
                try
                {
                    message = "Успешно";
                    return new XlParser(isDowload);
                }
                catch (Exception)
                {
                    message = "Ошибка";
                    return null;
                }
            }

            try
            {
                message = "Успешно";
                return new XlParser();
            }
            catch (Exception)
            {
                message = "Ошибка";
                return null;
            }

        }

        public void OpenTable()
        {
            

            XlParser table = LoadTable(out string message);

            //Excel.Workbook workBook = WorkExcel.Workbooks.Open(filepath);
            //Excel.Worksheet workSheet = (Excel.Worksheet)workBook.Sheets[1];

            List<Note> notes = ConvertTable(table.Workbook);

            foreach(Note note in notes)
            {
                Sheet.Add(note);
            }

        }

        public void SaveTable()
        {
            XlParser table = LoadTable(out string message);
            
            using (SaveFileDialog exportFile = new SaveFileDialog())
            {
                exportFile.Title = "Экспорт файла";
                exportFile.Filter = "Microsoft Office Excel Workbook(*.xlsx)|*.xlsx";

                if(DialogResult.OK == exportFile.ShowDialog())
                {
                    string filepath = exportFile.FileName;
                    table.Workbook.SaveAs(filepath);
                }
            }
        }

        private XLWorkbook WorksheetFill(XLWorkbook workbook)
        {
            var worksheet = workbook.Worksheets.Add("Список угроз");

            worksheet.Cell("A1").Value = "Общая информация";
            worksheet.Cell("F1").Value = "Последствия";
            worksheet.Cell("I1").Value = "Дополнительно";

            var rngtable1 = worksheet.Range("A1:E1");
            rngtable1.Merge();

            var rngtable2 = worksheet.Range("F1:H1");
            rngtable2.Merge();

            var rngtable3 = worksheet.Range("I1:J1");
            rngtable3.Merge();


            worksheet.Cell("A2").Value = "Идентификатор УБИ";
            worksheet.Cell("B2").Value = "Наименование УБИ";
            worksheet.Cell("C2").Value = "Описание";
            worksheet.Cell("D2").Value = "Источник угрозы (характеристика и потенциал нарушителя)";
            worksheet.Cell("E2").Value = "Объект воздействия";
            worksheet.Cell("F2").Value = "Нарушение конфиденциальности";
            worksheet.Cell("G2").Value = "Нарушение целостности";
            worksheet.Cell("H2").Value = "Нарушение доступности";
            worksheet.Cell("I2").Value = "Дата включения угрозы в БнД УБИ";
            worksheet.Cell("J2").Value = "Дата последнего изменения данных";

            int index = 3;
            foreach(Note note in Sheet)
            {
                worksheet.Cell($"A{index.ToString()}").Value = note.Id.ToString();
                worksheet.Cell($"B{index}").Value = note.Name;
                worksheet.Cell($"C{index}").Value = note.Description.ToString();
                worksheet.Cell($"D{index}").Value = note.Source;
                worksheet.Cell($"E{index}").Value = note.Threat;
                worksheet.Cell($"F{index}").Value = note.IsNotConfidential;
                worksheet.Cell($"G{index}").Value = note.IsComplete;
                worksheet.Cell($"H{index}").Value = note.IsAccessible;
                worksheet.Cell($"I{index}").Value = note.DateIn;
                worksheet.Cell($"J{index}").Value = note.DateRewrite;
            }

            return workbook;

        }

        public string UpdateTable()
        {
            
            XlParser table = LoadTable(out string message, true);

            if (table == null)
            {
                return message;
            }
            
            List<Note> notes = ConvertTable(table.Workbook);

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

            return message;
        }

        private List<Note> ConvertTable(XLWorkbook workbook)
        {
            var worksheet = workbook.Worksheet(1);
            var rows = worksheet.RangeUsed().RowsUsed();

            List<Note> list = new List<Note>();


            int skipLines = 0;
            foreach(var row in rows)
            {
                if (skipLines < 2)
                {
                    skipLines++;
                    continue;
                }
                try
                {
                    int id = Convert.ToInt32(row.Cell(1).Value.ToString());
                    string idInfo = $"УБИ.{id.ToString()}";
                    string name = row.Cell(2).Value.ToString();
                    string description = row.Cell(3).Value.ToString();
                    string source = row.Cell(4).Value.ToString();
                    string threat = row.Cell(5).Value.ToString();
                    string isNotConfidential = ParseInt(row.Cell(6).Value.ToString());
                    string isComplete = ParseInt(row.Cell(7).Value.ToString());
                    string isAccessible = ParseInt(row.Cell(8).Value.ToString());
                    string dateIn = row.Cell(9).Value.ToString();
                    string dateRewrite = row.Cell(10).Value.ToString();
                    list.Add(new Note(id, idInfo, name, description, source, threat, isNotConfidential, isComplete, isAccessible, dateIn, dateRewrite));
                }
                catch (Exception)
                {
                    continue;
                }
            }
            return list;
        }

        private string ParseInt(string parseItem)
        {
            bool success = Int32.TryParse(parseItem, out int value);

            if (success)
            {
                return value == 1 ? "Да" : "Нет";
            }

            return parseItem;
        }
    }
}
