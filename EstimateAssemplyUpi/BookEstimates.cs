using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Math;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.Win32;
using Application = Microsoft.Office.Interop.Excel.Application;
using Path = System.IO.Path;
using Shape = Microsoft.Office.Interop.Excel.Shape;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;
using XlHAlign = Microsoft.Office.Interop.Excel.XlHAlign;
using PrinterSettings = System.Drawing.Printing.PrinterSettings;

namespace EstimatesAssembly {
    class BookEstimates {
        private const string pageContent = @"\contentpage.xlsx";
        private const string pageTitle = @"\titlepage.xlsx";
        private const string pageResolution = @"\resolutionpage.xlsx";
        private const int PixelW = 50;
        private const int PixelH = 25;

        struct Ogl {
            public string col1;
            public string col2;
        }

        private string _nameBook;
        private string _pathBook;
        public Application Ex;
        public Workbook Wb;
        public Workbook TmpWb;
        // Для перенумерации листов книги
        private const int StopPos = 49;
        private const int EndPos = 59;
        private int delta = 0;

        public string NameBook {
            get { return _nameBook; }
            set { _nameBook = value; }
        }

        public string PathBook {
            get { return _pathBook; }
            set { _pathBook = value; }
        }

        private string printer = @"Microsoft Print to PDF";
        private string port = String.Empty;

        public BookEstimates() {
            Ex = new Application { Visible = false, DisplayAlerts = false };
        }

        public void SetActivePrinterPDF() {
            RegistryKey key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Devices");
            if (key != null) {
                object value = key.GetValue(printer);
                if (value != null) {
                    string[] values = value.ToString().Split(',');
                    if (values.Length >= 2) {
                        port = values[1];
                    }
                }
            }
            if (!Ex.ActivePrinter.StartsWith(printer)) {
                var split = Ex.ActivePrinter.Split(' ');
                if (split.Length >= 3) {
                    var prn = String.Format("{0} ({1})", printer,  port);
                    Ex.ActivePrinter = prn;
                }
                //MessageBox.Show("Test2", "Test2", MessageBoxButtons.OK);
            }
        }
        public void ShowExcel(Boolean show) {
            Ex.Visible = show;
        }
        // Тип сметы
        public static int FindTypeSheet(Worksheet sheet) {
            if (sheet.Cells.Find(@"ЛОКАЛЬНЫЙ СМЕТНЫЙ РАСЧЕТ №") != null) { return 1; }
            if (sheet.Cells.Find(@"ОБЪЕКТНЫЙ СМЕТНЫЙ РАСЧЕТ №") != null) { return 2; }
            if (sheet.Cells.Find(@"ВЕДОМОСТЬ РЕСУРСОВ") != null) { return 3; }
            if (sheet.Cells.Find(@"ВЕДОМОСТЬ ОБЪЕМОВ РАБОТ №") != null) { return 4; }
            if (sheet.Cells.Find(@"СВОДНЫЙ СМЕТНЫЙ РАСЧЕТ СТОИМОСТИ СТРОИТЕЛЬСТВА") != null) { return 5; }
            if (sheet.Cells.Find(@"Локальный ресурсный сметный расчет") != null) { return 6; }
            return 0;
        }

        // Добавить элемент(ы) в книгу
        public void AddSheetNew(string[] selectedItems) {
            if (selectedItems.Length == 0) {
                MessageBox.Show(@"Не выбрано ни одной сметы!", @"Внимание!");
                return;
            }
            Wb = Ex.Workbooks.Count == 0 ? Ex.Workbooks.Add() : Ex.ActiveWorkbook;
            Boolean isBookEstimate = false;
            string tmpfile = null;
            foreach (string selectedItem in selectedItems) {
                TmpWb = Ex.Workbooks.Open(selectedItem);
                Worksheet title = TmpWb.Sheets[1];
                if (title.Name.Equals(@"Титул")) {
                    isBookEstimate = true;
                    tmpfile = selectedItem;
                    break;
                }
            }
            if (isBookEstimate) {
                TmpWb = Ex.Workbooks.Open(tmpfile);
                foreach (Worksheet sheet in TmpWb.Sheets) {
                    sheet.Copy(Type.Missing, Wb.ActiveSheet);
                }
                TmpWb.Close();
            } else {
                foreach (string selectedItem in selectedItems) {
                    TmpWb = Ex.Workbooks.Open(selectedItem);
                    foreach (Worksheet sheet in TmpWb.Sheets) {
                        switch (FindTypeSheet(sheet)) {
                            case 1:
                                WorkWithExcelLs(sheet);
                                break;
                            case 2:
                                WorkWithExcelOs(sheet);
                                break;
                            case 3:
                                WorkWithExcelR(sheet);
                                break;
                            case 4:
                                WorkWithExcelVR(sheet);
                                break;
                            case 5:
                                WorkWithExcelSSR(sheet);
                                break;
                            case 6:
                                WorkWithExcelLRS(sheet);
                                break;
                        }
                        sheet.Copy(Type.Missing, Wb.ActiveSheet);
                    }
                    TmpWb.Close();
                }
            }
            foreach (string myvar in GetListSheet()) {
                if (myvar.Contains("Лист")) {
                    Wb.Sheets[myvar].Delete();
                }
            }
            SetActivePrinterPDF();
        }

        // Обработка сводного сметного расчета
        private void WorkWithExcelSSR(Worksheet sheet) {
            sheet.UsedRange.Font.Name = "Times New Roman";
            sheet.Range["A1:G5"].Clear();
            sheet.Range["B11"].Clear();
            sheet.Range["A15:H15"].Clear();
            sheet.Range["A13:H13"].Merge();
            sheet.Range["A15:H15"].Merge();
            sheet.Range["A16:H16"].Merge();
            sheet.Range["A15"].Value2 = MainFormAsm.iniSet.TbNameBuilding;
            sheet.Range["A15"].Font.Bold = true;
            sheet.Range["A15"].Font.Underline = true;
            sheet.Range["A15"].Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            Range find = sheet.Cells.Find("Итого \"Налоги и обязательные платежи\"");
            if (find != null) {
                Range price = sheet.Range["B18"];
                price.Value2 = @"Составлена в ценах по состоянию на " + QuarterFromDate(MainFormAsm.iniSet.CbPriceLevelO);
                sheet.Name = @"СС00";
            } else {
                Range price = sheet.Range["B18"];
                price.Value2 = @"Составлена в ценах по состоянию на 01.01.2000";
                sheet.Name = @"СС01";
            }

            // Подписи в конце страницы =======================================================
            if (MainFormAsm.iniSet.CbInsertSignOE) {
                // Всего по объектной смете
                Range findEnd = sheet.Cells.Find(@"Всего по сводному расчету");
                int rowEnd = findEnd.Row + 1;
                Range www = sheet.Range["A" + rowEnd.ToString() + ":J" + ((int)(rowEnd + 15)).ToString()];
                www.UnMerge();
                www.Clear();
                var rowGip = rowEnd + 3;
                var rowBoss = rowEnd + 6;
                var rowMadeIn = rowEnd + 9;
                // вставим надписи и ФИО
                Range rangeWork = sheet.Cells[rowGip, "C"];
                rangeWork.Value2 = @"Главный инженер проекта :";
                rangeWork.HorizontalAlignment = XlHAlign.xlHAlignRight;
                rangeWork = sheet.Cells[rowBoss, "C"];
                rangeWork.Value2 = @"Руководитель группы смет :";
                rangeWork.HorizontalAlignment = XlHAlign.xlHAlignRight;
                rangeWork = sheet.Cells[rowMadeIn, "C"];
                rangeWork.Value2 = @"Составил :";
                rangeWork.HorizontalAlignment = XlHAlign.xlHAlignRight;

                rangeWork = sheet.Cells[rowGip, "D"];
                rangeWork.Value2 = "_____________________________" + MainFormAsm.iniSet.CbGipText;
                rangeWork.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                rangeWork = sheet.Cells[rowBoss, "D"];
                rangeWork.Value2 = "_____________________________" + MainFormAsm.iniSet.TbHeadDepartment;
                rangeWork.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                rangeWork = sheet.Cells[rowMadeIn, "D"];
                rangeWork.Value2 = "_____________________________" + MainFormAsm.iniSet.CbMadeInText;
                rangeWork.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                // Вставим картинки
                if (MainFormAsm.iniSet.CbInsertSignSS) {
                    if (!MainFormAsm.iniSet.CbGip.Equals("")) {
                        InsertImage(ref sheet, rowGip, 5, MainFormAsm.iniSet.CbGipText);
                    }
                    if (!MainFormAsm.iniSet.TbHeadDepartment.Equals("")) {
                        InsertImage(ref sheet, rowBoss, 5, MainFormAsm.iniSet.TbHeadDepartment);
                    }
                    if (!MainFormAsm.iniSet.CbMadeIn.Equals("")) {
                        InsertImage(ref sheet, rowMadeIn, 5, MainFormAsm.iniSet.CbMadeInText);
                    }
                }
            }

        }

        // Обработка локального ресурсного сметного расчета
        private void WorkWithExcelLRS(Worksheet sheet) {
            sheet.UsedRange.Font.Name = "Times New Roman";
            Range find = sheet.Cells.Find("к Локальной смете №");
            string number = find.Value2;
            number = number.Substring(number.IndexOf("№") + 2);
            sheet.Name = "РС" + number;

            // Поработаем с подписями
            string firstName = "";
            string secondName = "";
            int stroka1 = 0;
            int stroka2 = 0;

            // Вначале все очистим от старых
            //int lastRow = FindLastRow(sheet);
            //sheet.Range["A" + (lastRow - 3).ToString(), "A" + lastRow].Clear();

            var range10 = sheet.Cells.Find(@"Составил");
            if (range10 != null) {
                stroka1 = range10.Row;
                var s1 = range10.Value2 as string;
                if (s1 != null) {
                    firstName = s1.Remove(0, s1.LastIndexOf('_') + 1).TrimEnd('\r', '\n', ' ');
                } // первое имя
            }
            var range20 = sheet.Cells.Find(@"Проверил");
            if (range20 != null) {
                stroka2 = range20.Row;
                var s2 = range20.Value2 as string;
                if (s2 != null) {
                    secondName = s2.Remove(0, s2.LastIndexOf('_') + 1).TrimEnd('\r', '\n', ' ');
                } // второе имя
            }
            // Очищаем и развоплощаем объединенные ячейки с подписями
            if (stroka1 != 0 && stroka2 != 0) {
                range20 = sheet.Range[sheet.Cells[stroka1, "A"], sheet.Cells[stroka2 + 1, "A"]];
                range20.Value2 = "";
                range20.UnMerge();
                sheet.Range[sheet.Cells[stroka1, "B"], sheet.Cells[stroka2, "Q"]].WrapText = false;
                sheet.Cells[stroka1, "D"] = @"Составил :";
                sheet.Cells[stroka1, "D"].HorizontalAlignment = XlHAlign.xlHAlignRight;
                sheet.Cells[stroka1, "E"] = "_____________________________" + firstName;
                sheet.Cells[stroka2, "D"] = @"Проверил :";
                sheet.Cells[stroka2, "D"].HorizontalAlignment = XlHAlign.xlHAlignRight;
                sheet.Cells[stroka2, "E"] = "_____________________________" + secondName;
            } else if (stroka1 == 0 && stroka2 != 0) {
                range20 = sheet.Range[sheet.Cells[stroka2, "A"], sheet.Cells[stroka2 + 1, "A"]];
                range20.Value2 = "";
                range20.UnMerge();
                sheet.Range[sheet.Cells[stroka2, "B"], sheet.Cells[stroka2, "Q"]].WrapText = false;
                sheet.Cells[stroka2, "D"] = @"Проверил :";
                sheet.Cells[stroka2, "D"].HorizontalAlignment = XlHAlign.xlHAlignRight;
                sheet.Cells[stroka2, "E"] = "_____________________________" + secondName;
            } else if (stroka1 != 0 && stroka2 == 0) {
                range20 = sheet.Range[sheet.Cells[stroka1, "A"], sheet.Cells[stroka1 + 1, "A"]];
                range20.Value2 = "";
                range20.UnMerge();
                sheet.Range[sheet.Cells[stroka1, "B"], sheet.Cells[stroka1, "Q"]].WrapText = false;
                sheet.Cells[stroka1, "D"] = @"Составил :";
                sheet.Cells[stroka1, "D"].HorizontalAlignment = XlHAlign.xlHAlignRight;
                sheet.Cells[stroka1, "E"] = "_____________________________" + firstName;
            }
            // Вставим подписи в ЛРС если нужно
            // Подписи в конце страницы =======================================================
            if (MainFormAsm.iniSet.CbInsertSignLR) {
                // вставим надписи и ФИО
                if (!firstName.Equals("") && stroka1 != 0) {
                    InsertImage(ref sheet, stroka1, 5, firstName);
                }
                if (!secondName.Equals("") && stroka2 != 0) {
                    InsertImage(ref sheet, stroka2, 5, secondName);
                }
            }

        }

        // Удалить элемент(ы) из книги
        public void DeleteSheet(ListView.SelectedListViewItemCollection selectedItems) {
            if (selectedItems.Count == 0) {
                MessageBox.Show(@"Не выбрано ни одно сметы!", @"Внимание!");
                return;
            }
            if (Ex.Workbooks.Count == 0) {
                return;
            }
            Ex.DisplayAlerts = false;
            foreach (ListViewItem selectedItem in selectedItems) {
                Worksheet worksheet = Wb.Sheets[selectedItem.Text];
                if (worksheet.Visible == XlSheetVisibility.xlSheetHidden) {
                    worksheet.Visible = XlSheetVisibility.xlSheetVisible;
                }
                if (Wb.Sheets.Count == 1) {
                    Wb.Sheets.Add();
                }
                Wb.Sheets[selectedItem.Text].Delete();
            }
        }

        // Возвращает список листов в книге
        public IEnumerable<string> GetListSheet() {
            var list = new List<string>();
            if (Ex.Workbooks.Count == 0) {
                return null;
            }
            Workbook workbook = Ex.ActiveWorkbook;
            if (workbook.Sheets.Count == 0) {
                return null;
            }
            for (int i = 1; i < workbook.Sheets.Count + 1; i++) {
                list.Add(workbook.Sheets[i].Name);
            }
            return list;
        }

        // Сохранение тома
        public void SaveWorkbook() {
            string fullname = Path.Combine(_pathBook, _nameBook + @".xls");
            if (File.Exists(fullname)) {
                DialogResult dlgres = MessageBox.Show(@"Книга уже существует. Переписать?", @"Внимание!",
                    MessageBoxButtons.OKCancel);
                if (dlgres == DialogResult.Cancel) {
                    return;
                }
            }
            Ex.DisplayAlerts = false;
            Ex.UserControl = true;
            try {
                Ex.ActiveWorkbook.SaveAs(fullname,
                    XlFileFormat.xlWorkbookNormal, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                MessageBox.Show(@"Книга успешно сохранена!");
            } catch (Exception e) {
                MessageBox.Show(e.Message);
            }
        }

        // Закрытие рабочего Excel-приложения
        public void CloseBook() {
            Ex.DisplayAlerts = false;
            Ex.UserControl = true;
            Ex.Quit();
        }

        // Инициализация книги
        public void initBook(string bookfile) {
            if (File.Exists(bookfile)) {
                Wb = Ex.Workbooks.Open(bookfile);
                Ex.DisplayAlerts = false;
                Ex.UserControl = true;
            }
        }

        // Перемещение элемента вверх по списку
        public void MoveUpsheet(ListView.SelectedListViewItemCollection selectedItems) {
            if (selectedItems.Count == 0) {
                MessageBox.Show(@"Не выбрано ни одно сметы!", @"Внимание!");
                return;
            }
            if (Ex.Workbooks.Count == 0) {
                return;
            }
            foreach (ListViewItem selectedItem in selectedItems) {
                Worksheet worksheet = Wb.Sheets[selectedItem.Text];
                int i = worksheet.Index;
                if (i > 1) {
                    worksheet.Move(Wb.Sheets[i - 1], Type.Missing);
                }
            }
        }

        // Перемещение элемента вниз по списку
        public void MoveDownsheet(ListView.SelectedListViewItemCollection selectedItems) {
            if (selectedItems.Count == 0) {
                MessageBox.Show(@"Не выбрано ни одно сметы!", @"Внимание!");
                return;
            }
            if (Ex.Workbooks.Count == 0) {
                return;
            }
            foreach (ListViewItem selectedItem in selectedItems) {
                Worksheet worksheet = Wb.Sheets[selectedItem.Text];
                int i = worksheet.Index;
                if (i < Wb.Sheets.Count) {
                    worksheet.Move(Type.Missing, Wb.Sheets[i + 1]);
                }
            }
        }

        // Сортировка списка согласно правилам сметчиков
        public void SortWorksheets() {
            List<string> list = new List<string>();
            if (Ex.ActiveWorkbook == null) {
                return;
            }
            foreach (Worksheet ws in Ex.ActiveWorkbook.Sheets) {
                list.Add(ws.Name);
            }
            list.Sort(Compare);
            Workbook wb = Ex.ActiveWorkbook;
            foreach (string str in list) {
                Worksheet ws = wb.Sheets[str];
                ws.Move(Wb.Sheets[list.IndexOf(str) + 1], Type.Missing);
            }
        }

        // Компаратор для сортировщика
        public int Compare(String x, String y) {
            Regex pattern = new Regex(@"(?:[ЛОРСЕ]+?[С]?)(?<ss>(\(?(\d{2,4})[-|\.])*(\d{1,4})(\)?))");
            MatchCollection mc1 = pattern.Matches(x);
            MatchCollection mc2 = pattern.Matches(y);
            if (mc1.Count == 0) {
                x = @"О0";
            }
            if (mc2.Count == 0) {
                y = @"О0";
            }
            int compareResult = 0;
            Int64 xx1;
            Int64 yy1;
            int xi = x.IndexOf(".");
            int yi = y.IndexOf(".");
            if (xi > 0) {
                x = x.Remove(xi);
            }
            if (yi > 0) {
                y = y.Remove(yi);
            }
            if (TwoChar(x)) {
                xx1 = Int64.Parse(x.Substring(2).Replace("-", "").Replace(".", "").Replace("(", "").Replace(" ", "").Replace(")", "").PadRight(14, '0'));
            } else {
                xx1 = Int64.Parse(x.Substring(1).Replace("-", "").Replace(".", "").Replace("(", "").Replace(" ", "").Replace(")", "").PadRight(14, '0'));
            }

            if (TwoChar(y)) {
                yy1 = Int64.Parse(y.Substring(2).Replace("-", "").Replace(".", "").Replace("(", "").Replace(" ", "").Replace(")", "").PadRight(14, '0'));
            } else {
                yy1 = Int64.Parse(y.Substring(1).Replace("-", "").Replace(".", "").Replace("(", "").Replace(" ", "").Replace(")", "").PadRight(14, '0'));
            }

            Int64 xx2 = ConvertChar(x);
            Int64 yy2 = ConvertChar(y);

            if (xx1 > yy1) {
                compareResult = 1;
            } else if (xx1 < yy1) {
                compareResult = -1;
            } else if (xx1 == yy1) {
                if (xx2 > yy2) {
                    compareResult = 1;
                } else if (xx2 < yy2) {
                    compareResult = -1;
                } else {
                    compareResult = 0;
                }
            }
            return compareResult;
        }

        // В названии первых символов - два?
        private bool TwoChar(string s) {
            if (s.Contains("С")) {
                return true;
            }
            return false;
        }

        // Выдает число в зависимости от первых двух символов наименования
        private Int64 ConvertChar(string a) {
            switch (a.Substring(0, 1)) {
                case "ОС":
                case "О":
                    return 1;
                case "ЛC":
                case "Л":
                    return 2;
                case "Р":
                    return 3;
                case "СС":
                    return 4;
                case "РС":
                    return 5;
                default:
                    return 0;
            }
        }

        // Дополнительная обработка таблиц
        public void AdaptionSheets() {
            Workbook mainBook = Ex.ActiveWorkbook;
            Range r;
            if (mainBook == null) {
                return;
            }
            foreach (Worksheet worksheet in mainBook.Sheets) {
                if (!worksheet.Name.Equals(@"Титул")
                    && !worksheet.Name.Equals(@"Оглавление")
                    && !worksheet.Name.Equals(@"Разрешение")) {
                    worksheet.Activate();
                    HPageBreaks hbreak = worksheet.HPageBreaks;
                    int pageCount = hbreak.Count - 1;
                    if (pageCount != 0) {
                        var a = hbreak.Item[pageCount];
                        r = hbreak.Item[pageCount].Location;
                        int t = FindLastRow(worksheet);
                        int t1 = r.Row;
                        if ((t - t1) < 12 && (t - t1) > 0) {
                            var tmpr = worksheet.Range["A" + Convert.ToString(r.Row - 5)];
                            hbreak.Item[pageCount].Location = tmpr;
                        }
                    }
                }
            }
        }

        private int FindLastRow(Worksheet worksheet) {
            //            Range last = worksheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            //            return last.Row;
            int lastUsedRow = 1;
            Range range = worksheet.UsedRange;
            for (int i = 1; i < range.Columns.Count; i++) {
                int lastRow = range.Rows.Count;
                for (int j = range.Rows.Count; j > 0; j--) {
                    if (lastUsedRow < lastRow) {
                        lastRow = j;
                        if (!String.IsNullOrWhiteSpace(Convert.ToString((worksheet.Cells[j, i] as Range).Value))) {
                            if (lastUsedRow < lastRow)
                                lastUsedRow = lastRow;
                            if (lastUsedRow == range.Rows.Count)
                                return lastUsedRow - 1;
                            break;
                        }
                    } else
                        break;
                }
            }
            return lastUsedRow;
        }

        public void NumberingPage() {
            Workbook mainBook = Ex.ActiveWorkbook;
            if (mainBook == null) {
                return;
            }
            foreach (Worksheet worksheet in mainBook.Sheets) {
                if (worksheet.Name.Contains(@"Титул"))
                    worksheet.Delete();
                else
                    if (worksheet.Name.Contains(@"Разрешение"))
                    worksheet.Delete();
                else
                        if (worksheet.Name.Contains(@"Оглавление"))
                    worksheet.Delete();
                else
                            if (worksheet.Name.Contains(@"Лист"))
                    worksheet.Delete();
            }
            // Вставим оглавление
            Workbook tmpContent = Ex.Workbooks.Open(MainFormAsm.iniSet.TxtToolsFilesPath + @"\contentpage.xlsx");
            tmpContent.Worksheets[1].Copy(mainBook.Sheets[1], Type.Missing);
            tmpContent.Close();
            Worksheet ogl = mainBook.Sheets[1];
            //Worksheet title = mainBook.Sheets[1];
            ogl.Name = @"Оглавление";
            ogl.Cells[2, 5] = _nameBook;
            ogl.PageSetup.Zoom = false;

            Worksheet title = mainBook.Sheets.Add(mainBook.Sheets.Item[1], Type.Missing, 1, XlSheetType.xlWorksheet);
            title.Name = @"Титул";
            title.PageSetup.Zoom = false;
            
            // Включим разрывы страниц
            foreach (Worksheet worksheet in mainBook.Sheets) {
                worksheet.Select();
                Ex.ActiveWindow.View = XlWindowView.xlPageBreakPreview;
            }

            foreach (Worksheet worksheet in mainBook.Sheets) {
                worksheet.Select();
                worksheet.PageSetup.Zoom = false;
                worksheet.PageSetup.FitToPagesWide = 1;
                worksheet.PageSetup.FitToPagesTall = 999;
                if (worksheet.VPageBreaks.Count > 0) {
                    worksheet.VPageBreaks.get_Item(1).DragOff(XlDirection.xlToRight, 1);
                }
            }
            int ns = 3;
            int x = 1;
            Ogl a = new Ogl();
            if (mainBook.Sheets.Count < StopPos - 1) {
                Range clr = ogl.Range["A60", "L125"];
                clr.Clear();
            }
            foreach (Worksheet worksheet in mainBook.Sheets) {
                worksheet.Select();
                worksheet.PageSetup.FirstPageNumber = x;
                worksheet.PageSetup.RightFooter = "&P";
                worksheet.PageSetup.LeftHeader = " ";
                worksheet.PageSetup.CenterHeader = " ";
                worksheet.PageSetup.RightHeader = " ";
                a = GetColumnsSheet(worksheet);
                if (!worksheet.Name.Equals("Титул") && !worksheet.Name.Equals("Оглавление")) {
                    ogl.Cells[ns, 4] = ns - delta - 2;
                    ogl.Cells[ns, 5] = a.col1;
                    ogl.Cells[ns, 8] = a.col2;
                    Range range_1 = ogl.Cells[ns, 12];
                    range_1.Value2 = String.Format("{0}", worksheet.PageSetup.FirstPageNumber);
                    ogl.Hyperlinks.Add(range_1, "", "'" + worksheet.Name + "'!A1", Type.Missing, "Hyperlink Test");
                    Range ssss = ogl.Rows[ns];
                    ssss.RowHeight = 12.75;
                    ns++;
                }
                if (ns == StopPos + 1) {
                    ns = EndPos + 1;
                    delta = 10;
                }
                x = worksheet.PageSetup.FirstPageNumber + worksheet.PageSetup.Pages.Count;
                worksheet.PageSetup.Zoom = 100;
            }
            delta = 0;
            title.Delete();
            // Вставим титульные листы
            Workbook tmpTitle = Ex.Workbooks.Open(MainFormAsm.iniSet.TxtToolsFilesPath + pageTitle);
            tmpTitle.Worksheets[1].Copy(mainBook.Sheets[1], Type.Missing);
            tmpTitle.Close();
            Worksheet titles = mainBook.Sheets[1];
            titles.Name = @"Титул";
            TitleFill(ref titles);
            //AdaptionSheets(ref pgBar);

            if (int.Parse(MainFormAsm.iniSet.NumModification) != 0) {
                Workbook tmpResolution = Ex.Workbooks.Open(MainFormAsm.iniSet.TxtToolsFilesPath + @"\resolutionpage.xlsx");
                tmpResolution.Worksheets[1].Copy(mainBook.Sheets[2], Type.Missing);
                tmpResolution.Close();
                Worksheet resolution = mainBook.Sheets[2];
                resolution.Name = @"Разрешение";
                ResolutionFill(ref resolution);
            }
            if (ns < EndPos + 1) {
                StampFill(false, ref ogl, x - 2);
            } else {
                StampFill(true, ref ogl, x - 2);
            }
            
        }

        // Заполним Разрешение
        private void ResolutionFill(ref Worksheet resolution) {
            resolution.UsedRange.Font.Name = "Times New Roman";
            resolution.Cells[9, 5] = MainFormAsm.iniSet.NumModification;
            resolution.Cells[9, 7] = MainFormAsm.iniSet.TbPageNumber;
            resolution.Cells[46, 8] = MainFormAsm.iniSet.TbChiefEngineer;
            resolution.Cells[47, 8] = MainFormAsm.iniSet.CbGipText;
            resolution.Cells[48, 8] = MainFormAsm.iniSet.CbMadeInText;
            resolution.Cells[49, 8] = MainFormAsm.iniSet.CbMadeInText;
            resolution.Cells[46, 13] = MainFormAsm.iniSet.DateAjustment.ToString("MM.yy", CultureInfo.CreateSpecificCulture("ru-RU"));
            resolution.Cells[47, 13] = MainFormAsm.iniSet.DateAjustment.ToString("MM.yy", CultureInfo.CreateSpecificCulture("ru-RU"));
            resolution.Cells[48, 13] = MainFormAsm.iniSet.DateAjustment.ToString("MM.yy", CultureInfo.CreateSpecificCulture("ru-RU"));
            resolution.Cells[49, 13] = MainFormAsm.iniSet.DateAjustment.ToString("MM.yy", CultureInfo.CreateSpecificCulture("ru-RU"));
            InsertImage(ref resolution, 47, 11, MainFormAsm.iniSet.TbChiefEngineer);
            InsertImage(ref resolution, 48, 11, MainFormAsm.iniSet.CbGipText);
            InsertImage(ref resolution, 49, 11, MainFormAsm.iniSet.CbMadeInText);
            InsertImage(ref resolution, 50, 11, MainFormAsm.iniSet.CbMadeInText);
            resolution.Cells[46, 15] = "ООО \"ИПИГАЗ\"";
            resolution.Cells[48, 23] = "1";
            resolution.Cells[3, 5] = MainFormAsm.iniSet.TbDocumentNumber;
            // 
            String loverStr = MainFormAsm.iniSet.ListTypeDocument.ToLower();
            String volNum = MainFormAsm.iniSet.NumVolumeNumber;
            String bookNum = MainFormAsm.iniSet.NumBookNumber;
            String partNum = MainFormAsm.iniSet.NumPartNumber;
            String lStr = loverStr.Substring(0, 1).ToUpper() + loverStr.Substring(1, loverStr.Length - 1);
            String str = @"Инв.№" + MainFormAsm.iniSet.TbInventoryNumber + "\n" +
                MainFormAsm.iniSet.TbCodeObject + "\n" +
                @"Том " + volNum + "." + bookNum + "." + partNum + " \"" + lStr + "\"";
            resolution.Cells[1, 9] = str;
            //
            resolution.Cells[1, 19] = MainFormAsm.iniSet.TbNameBuilding;
        }

        // Заполним титул
        private void TitleFill(ref Worksheet title) {
            title.UsedRange.Font.Name = "Times New Roman";
            //            title.Cells[8, 3] = MainFormAsm.iniSet.TbCertificate; // Свидетельство
            //            title.Cells[10, 3] = MainFormAsm.iniSet.TbCustomer; // Заказчик
            title.Cells[14, 2] = MainFormAsm.iniSet.TbNameBuilding;
            title.Cells[19, 2] = MainFormAsm.iniSet.TbNameObject;
            title.Cells[25, 2] = MainFormAsm.iniSet.CbStageDevelope;
            title.Cells[30, 2] = MainFormAsm.iniSet.ListTypeDocument;
            title.Cells[32, 2] = MainFormAsm.iniSet.TbCodeObject;

//            if (!MainFormAsm.iniSet.CbRebuild) {
//                title.Cells[22, 3] = "РАЗДЕЛ " + int.Parse(MainFormAsm.iniSet.NumVolumeNumber) + " \"СМЕТА НА СТРОИТЕЛЬСТВО\"";
//                title.Cells[70, 3] = "РАЗДЕЛ " + int.Parse(MainFormAsm.iniSet.NumVolumeNumber) + " \"СМЕТА НА СТРОИТЕЛЬСТВО\"";
//            } else {
//                title.Cells[22, 3] = "РАЗДЕЛ " + int.Parse(MainFormAsm.iniSet.NumVolumeNumber) + " \"СМЕТА НА КАПИТАЛЬНЫЙ РЕМОНТ\"";
//                title.Cells[70, 3] = "РАЗДЕЛ " + int.Parse(MainFormAsm.iniSet.NumVolumeNumber) + " \"СМЕТА НА КАПИТАЛЬНЫЙ РЕМОНТ\"";
//            }

//            title.Cells[24, 3] = @"ЧАСТЬ 2 " + MainFormAsm.iniSet.ListTypeDocument.ToUpper();

//            title.Cells[25, 3] = "КНИГА " + MainFormAsm.iniSet.NumBookNumber;

//            string sss = "ТОМ " + MainFormAsm.iniSet.NumVolumeNumber + "." +
//                         MainFormAsm.iniSet.NumBookNumber + "." +
//                         MainFormAsm.iniSet.NumPartNumber;
//            title.Cells[29, 3] = sss;

            title.Cells[43, 25] = MainFormAsm.iniSet.TbChiefEngineer;
            InsertImage(ref title, 43, 18, MainFormAsm.iniSet.TbChiefEngineer);
            title.Cells[46, 25] = MainFormAsm.iniSet.CbGipText;
            InsertImage(ref title, 46, 18, MainFormAsm.iniSet.CbGipText);

            title.Cells[57, 17] = MainFormAsm.iniSet.TbYearTitle; // Год 

//            switch (int.Parse(MainFormAsm.iniSet.NumModification)) {
//                case 0:
//                    title.Range["D38", "G39"].Clear();
//                    title.Range["D87", "G88"].Clear();
//                    break;
//                default:
//                    title.Cells[39, 4] = int.Parse(MainFormAsm.iniSet.NumModification);
//                    title.Cells[39, 5] = MainFormAsm.iniSet.TbDocumentNumber;
//                    InsertImage(ref title, 39, 6, MainFormAsm.iniSet.CbMadeInText);
//                    title.Cells[39, 7] = MainFormAsm.iniSet.DateAjustment.ToString("\tMM/yyyy");
//                    title.Cells[88, 4] = int.Parse(MainFormAsm.iniSet.NumModification);
//                    title.Cells[88, 5] = MainFormAsm.iniSet.TbDocumentNumber;
//                    InsertImage(ref title, 88, 6, MainFormAsm.iniSet.CbMadeInText);
//                    title.Cells[88, 7] = MainFormAsm.iniSet.DateAjustment.ToString("\tMM/yyyy");
//                    break;
//            }
        }

        // Заполним штамп оглавления
        private void StampFill(Boolean twoPage, ref Worksheet stamp, int x) {
            stamp.Cells[EndPos - 4, 5] = MainFormAsm.iniSet.CbMadeInText;
            InsertImage(ref stamp, EndPos - 5, 7, MainFormAsm.iniSet.CbMadeInText);
            stamp.Cells[EndPos - 2, 5] = MainFormAsm.iniSet.TbHeadDepartment;
            InsertImage(ref stamp, EndPos - 3, 7, MainFormAsm.iniSet.TbHeadDepartment);
            stamp.Cells[EndPos - 1, 5] = MainFormAsm.iniSet.CbGipText;
            InsertImage(ref stamp, EndPos - 2, 7, MainFormAsm.iniSet.CbGipText);
            stamp.Cells[EndPos - 4, 8] = MainFormAsm.iniSet.DateToStamp.ToString("MM.yy");
            stamp.Cells[EndPos - 2, 8] = MainFormAsm.iniSet.DateToStamp.ToString("MM.yy");
            stamp.Cells[EndPos - 1, 8] = MainFormAsm.iniSet.DateToStamp.ToString("MM.yy");
            stamp.Cells[EndPos - 7, 9] = MainFormAsm.iniSet.TbCodeObject;
            stamp.Cells[EndPos - 5, 9] = MainFormAsm.iniSet.TbNameBuilding + "\nОбъектные и локальные сметы";
            stamp.Cells[EndPos - 2, 10] = "ООО «Югорский Проектный Институт»";
            stamp.Cells[EndPos - 4, 10] = MainFormAsm.iniSet.CbStageDevelope.Substring(0, 1);
            stamp.Cells[EndPos - 4, 11] = "1";
            stamp.Cells[EndPos - 4, 12] = (x + 1).ToString(CultureInfo.InvariantCulture);
            stamp.Cells[EndPos - 4, 2] = MainFormAsm.iniSet.TbInventoryNumber;
            if (int.Parse(MainFormAsm.iniSet.NumModification) != 0) {
                stamp.Cells[EndPos - 6, 3] = MainFormAsm.iniSet.NumModification;
                stamp.Cells[EndPos - 6, 4] = "-";
                stamp.Cells[EndPos - 6, 5] = MainFormAsm.iniSet.TbPageNumber;
                stamp.Cells[EndPos - 6, 6] = MainFormAsm.iniSet.TbDocumentNumber;
                stamp.Cells[EndPos - 6, 8] = MainFormAsm.iniSet.DateAjustment.ToString("MM.yy");
                InsertImage(ref stamp, EndPos - 7, 7, MainFormAsm.iniSet.CbMadeInText);
            }

            if (twoPage) {
                stamp.Cells[EndPos + 55, 2] = MainFormAsm.iniSet.TbInventoryNumber;
                stamp.Cells[EndPos + 57, 9] = MainFormAsm.iniSet.TbCodeObject;
            }
        }

        // Вставить картинку
        private void InsertImage(ref Worksheet sheet, int y, int x, string fio) {
            char[] charsToTrim = { '\n', '\r', ' ' };
            Shape shape = null;     
            Range range = sheet.Cells[y, x];
            fio = fio.TrimEnd(charsToTrim);
            float xx = (float)((double)range.Left);
            float yy = (float)((double)range.Top) - 4;
            try {
                var fName1 = MainFormAsm.iniSet.TxtImagePath + @"\" + ConvertName(fio) + ".jpg";
                var fName2 = MainFormAsm.iniSet.TxtImagePath + @"\" + ConvertName(fio) + ".tif";
                var fName3 = MainFormAsm.iniSet.TxtImagePath + @"\" + ConvertName(fio) + ".tiff";
                if (File.Exists(fName1)) {
                    shape = sheet.Shapes.AddPicture(fName1, MsoTriState.msoTrue, MsoTriState.msoTrue, xx, yy, PixelW, PixelH);
                } else if (File.Exists(fName2)) {
                    shape = sheet.Shapes.AddPicture(fName2, MsoTriState.msoTrue, MsoTriState.msoTrue, xx, yy, PixelW, PixelH);
                } else if (File.Exists(fName3)) {
                    shape = sheet.Shapes.AddPicture(fName3, MsoTriState.msoTrue, MsoTriState.msoTrue, xx, yy, PixelW, PixelH);
                }
                if (shape != null) {
                }
                if (shape != null) {
                    shape.PictureFormat.TransparentBackground = MsoTriState.msoTrue;
                    shape.PictureFormat.TransparencyColor = ColorTranslator.ToOle(Color.White);
                    shape.Fill.Visible = MsoTriState.msoFalse;
                }
            } catch (Exception e) {
                MessageBox.Show(e.Message, @"Ошибка при работе с изображением!");
            }
        }

        private static string ConvertName(string name) {
            string n = name.Replace(".", "_").Replace(" ", "_");
            n = n.Substring(0, n.Length - 1);
            return n;
        }

        // Вытаскиваем из таблицы номер и наименование сметы или объекта
        private Ogl GetColumnsSheet(_Worksheet worksheet) {
            Ogl o = new Ogl();
            Range range = worksheet.Cells.Find(@"ЛОКАЛЬНЫЙ СМЕТНЫЙ РАСЧЕТ");
            if (range != null) {
                string num = worksheet.Name.Substring(2);
                o.col1 = @"ЛСР №" + num;
                //o.col2 = @"локальный сметный расчет";
                o.col2 = worksheet.Range["A12"].Text;
                return o;
            }
            range = worksheet.Cells.Find(@"ОБЪЕКТНЫЙ СМЕТНЫЙ РАСЧЕТ");
            if (range != null) {
                string num = worksheet.Name.Substring(2);
                o.col1 = @"ОСР №" + num;
                o.col2 = @"объектный сметный расчет";
                return o;
            }
            range = worksheet.Cells.Find(@"ВЕДОМОСТЬ РЕСУРСОВ");
            if (range != null) {
                string num = worksheet.Name.Substring(1);
                o.col1 = @"РВ №" + num;
                o.col2 = @"Ресурсная ведомость";
                return o;
            }
            range = worksheet.Cells.Find(@"СВОДНЫЙ СМЕТНЫЙ РАСЧЕТ СТОИМОСТИ СТРОИТЕЛЬСТВА");
            if (range != null) {
                string tek = worksheet.Name.Substring(2);
                string price = worksheet.Range["B18"].Text;
                if (price != null) price = price.Substring(price.LastIndexOf(@"на"));
                if (tek.Contains("01")) {
                    tek = "баз";
                    o.col2 = @"Сводный сметный расчет в базовых ценах " + price;
                } else {
                    tek = "тек";
                    o.col2 = @"Сводный сметный расчет в текущих ценах " + price;
                }
                o.col1 = @"ССР " + tek;
                return o;
            }
            range = worksheet.Cells.Find(@"Локальный ресурсный сметный расчет");
            if (range != null) {
                string num = worksheet.Name.Substring(2);
                o.col1 = @"ЛРС №" + num;
                o.col2 = @"Локальный ресурсный сметный расчет";
                return o;
            }
            return o;
        }

        public void RebuildWorksheets() {
            Workbook mainBook = Ex.ActiveWorkbook;
            foreach (Worksheet worksheet in mainBook.Sheets) {
                string sss = WorkWithExcelLs(worksheet);
            }
        }

        private string QuarterFromDate(DateTime value) {
            int a = DateAndTime.DatePart(DateInterval.Quarter, value);
            int b = DateAndTime.DatePart(DateInterval.Year, value);
            if (MainFormAsm.iniSet.CbQuarter) {
                return String.Format("{0}-й квартал {1} года.", a, b);
            } else {
                return value.ToString("dd MMMM yyyy", CultureInfo.CreateSpecificCulture("ru-RU"));
            }
        }

        public static string RenameName(string name) {
            Regex pattern = new Regex(@"(?:\D*)(?<ss>((\d{2,4})[-|\.])*(\d{1,4}))");
            MatchCollection mc = pattern.Matches(name);
            if (mc.Count > 0) {
                GroupCollection groups = mc[0].Groups;
                return groups["ss"].Value;
            } else {
                return name;
            }
        }
        public static string RemoveBeginPos(string name) {
            Regex pattern = new Regex(@"(?:\D*)(?<ss>(\(?(\d{2,4})[-|\.])*(\d{1,4})(\)?))");
            MatchCollection mc = pattern.Matches(name);
            string num = null;
            if (mc.Count > 0) {
                GroupCollection groups = mc[0].Groups;
                num = groups["ss"].Value;
            }
            if (num != null) {
                int ii = name.IndexOf(num, System.StringComparison.Ordinal);
                var charEnd = name.Length;
                return name.Substring(num.Length + ii);
            }
            return name;
        }

        // Локальные сметы. Обработка
        private string WorkWithExcelLs(Worksheet sheet) {
            sheet.UsedRange.Font.Name = "Times New Roman";
            sheet.Range["A1:Q5"].Clear();
            string numberEstimate = sheet.Range["G9"].Text;
            string nameObject = MainFormAsm.iniSet.TbNameBuilding;
            numberEstimate = numberEstimate.Substring(numberEstimate.LastIndexOf("№") + 2);
            //    lLastCol = ActiveSheet.UsedRange.Column + ActiveSheet.UsedRange.Columns.Count - 1
            //int col = sheet.UsedRange.Column + sheet.UsedRange.Columns.Count - 1;
            Range tmprange = sheet.Range["A6:M6"];
            tmprange.Merge();
            tmprange.Font.Bold = true;
            tmprange.Font.Underline = true;
            tmprange.EntireRow.RowHeight = 20;
            tmprange.EntireRow.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            tmprange.Value2 = nameObject;

            tmprange = sheet.Range["A7:Q7"];
            tmprange.Merge();
            tmprange.Font.Italic = true;
            tmprange.Merge();
            tmprange.Value2 = "наименование стройки";
            string tmp = sheet.Range["D12"].Text;

            tmprange = sheet.Range["A12:Q12"];
            tmprange.Clear();
            tmprange.Merge();
            tmprange.Value2 = tmp;
            tmprange.Font.Name = "Times New Roman";
            tmprange.Font.Underline = true;
            tmprange.EntireRow.HorizontalAlignment = XlHAlign.xlHAlignCenter;

            //(наименование работ и затрат, наименование объекта)
            tmprange = sheet.Range["A13:Q13"];
            tmprange.Merge();
            // Это уровень цен ===============================================
            Range rangeWork = sheet.Range["D19"]; //rangeWork = "Составлен в ценах по состоянию на"
            if (rangeWork != null) {
                if (MainFormAsm.iniSet.CbQuarter) {
                    rangeWork.Value2 = @"Составлена в ценах по состоянию на " + QuarterFromDate(MainFormAsm.iniSet.CbPriceLevelO);
                    //                    rangeWork.HorizontalAlignment = HorizontalAlignment.Right;
                }
            }
            // Поработаем с подписями
            string firstName = "";
            string secondName = "";
            int stroka1 = 0;
            int stroka2 = 0;

            // Смета нового образца
            // Вначале все очистим от старых
            var range10 = sheet.Cells.Find(@"Составил");
            if (range10 != null) {
                stroka1 = range10.Row;
                var s1 = range10.Value2 as string;
                if (s1 != null) {
                    firstName = s1.Remove(0, s1.LastIndexOf('_') + 1).TrimEnd('\r', '\n', ' ');
                } // первое имя
            }
            var range20 = sheet.Cells.Find(@"Проверил");
            if (range20 != null) {
                stroka2 = range20.Row;
                var s2 = range20.Value2 as string;
                if (s2 != null) {
                    secondName = s2.Remove(0, s2.LastIndexOf('_') + 1).TrimEnd('\r', '\n', ' ');
                } // второе имя
            }
            // Очищаем и развоплощаем объединенные ячейки с подписями
            if (stroka1 != 0 && stroka2 != 0) {
                range20 = sheet.Range[sheet.Cells[stroka1, "A"], sheet.Cells[stroka2 + 1, "A"]];
                range20.Value2 = "";
                range20.UnMerge();
                sheet.Range[sheet.Cells[stroka1, "B"], sheet.Cells[stroka2, "Q"]].WrapText = false;
                sheet.Cells[stroka1, "D"] = @"Составил :";
                sheet.Cells[stroka1, "D"].HorizontalAlignment = XlHAlign.xlHAlignRight;
                sheet.Cells[stroka1, "E"] = "_____________________________" + firstName;
                sheet.Cells[stroka2, "D"] = @"Проверил :";
                sheet.Cells[stroka2, "D"].HorizontalAlignment = XlHAlign.xlHAlignRight;
                sheet.Cells[stroka2, "E"] = "_____________________________" + secondName;
            } else if (stroka1 == 0 && stroka2 != 0) {
                range20 = sheet.Range[sheet.Cells[stroka2, "A"], sheet.Cells[stroka2 + 1, "A"]];
                range20.Value2 = "";
                range20.UnMerge();
                sheet.Range[sheet.Cells[stroka2, "B"], sheet.Cells[stroka2, "Q"]].WrapText = false;
                sheet.Cells[stroka2, "D"] = @"Проверил :";
                sheet.Cells[stroka2, "D"].HorizontalAlignment = XlHAlign.xlHAlignRight;
                sheet.Cells[stroka2, "E"] = "_____________________________" + secondName;
            } else if (stroka1 != 0 && stroka2 == 0) {
                range20 = sheet.Range[sheet.Cells[stroka1, "A"], sheet.Cells[stroka1 + 1, "A"]];
                range20.Value2 = "";
                range20.UnMerge();
                sheet.Range[sheet.Cells[stroka1, "B"], sheet.Cells[stroka1, "Q"]].WrapText = false;
                sheet.Cells[stroka1, "D"] = @"Составил :";
                sheet.Cells[stroka1, "D"].HorizontalAlignment = XlHAlign.xlHAlignRight;
                sheet.Cells[stroka1, "E"] = "_____________________________" + firstName;
            }
            // Вставим подписи в ЛС если нужно
            // Подписи в конце страницы =======================================================
            if (MainFormAsm.iniSet.CbInsertSignLE) {
                // вставим надписи и ФИО
                if (!firstName.Equals("") && stroka1 != 0) {
                    InsertImage(ref sheet, stroka1, 6, firstName);
                }
                if (!secondName.Equals("") && stroka2 != 0) {
                    InsertImage(ref sheet, stroka2, 6, secondName);
                }
            }

            // Подписать страницу
            sheet.Name = @"ЛС" + numberEstimate;
            //SetRowHeigths(ref sheet, ref rangeWork);
            return numberEstimate;
        }

        // Объектные сметы. Обработка
        private string WorkWithExcelOs(Worksheet sheet) {
            sheet.UsedRange.Font.Name = "Times New Roman";
            // Затрем все лишнее сверху ===============================================
            ((Range)sheet.Range["J1"]).Clear();
            sheet.Range["A2:J2"].Merge();
            sheet.Range["A2:J2"].Font.Bold = true;
            sheet.Range["A2:J2"].Font.Underline = true;
            sheet.Range["A2:J2"].Value2 = MainFormAsm.iniSet.TbNameBuilding;
            string numberOE = ((Range)sheet.Range["G5"]).Text;
            sheet.Range["G5"].Clear();
            sheet.Range["E5"].Value2 += numberOE;
            sheet.Range["A5:J5"].Merge();
            sheet.Range["A5:J5"].Font.Underline = true;
            sheet.Range["A5:J5"].Font.Bold = true;

            string tmp = sheet.Range["D8"].Text;
            sheet.Range["A8:J8"].Clear();
            sheet.Range["A8:J8"].Merge();
            sheet.Range["A8:J8"].Value2 = tmp;
            sheet.Range["A8:J8"].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            sheet.Range["A8:J8"].Font.Name = "Times New Roman";
            sheet.Range["A8:J8"].Font.Underline = true;

            Range rangeWork;
            string nameWorks = numberOE;
            // Капремонт
            if (MainFormAsm.iniSet.CbRebuild) {
                rangeWork = sheet.Cells.Find(@"строительных работ");
                if (rangeWork != null) rangeWork.Value2 = @"ремонтно-строительных работ";
                rangeWork = sheet.Cells.Find(@"монтажных работ");
                if (rangeWork != null) rangeWork.Value2 = @"ремонтно-монтажных работ";
                rangeWork = sheet.Cells.Find(@"мебели, инвентаря");
                if (rangeWork != null) rangeWork.Value2 = @"комплектующих и запасных частей";
                rangeWork = sheet.Cells.Find(@"на строительство");
                if (rangeWork != null) rangeWork.Value2 = @"";
            }
            // Это уровень цен ===============================================
            rangeWork = sheet.Range["C14"];
            if (MainFormAsm.iniSet.CbQuarter) {
                rangeWork.Value2 = @"Составлена в ценах по состоянию на " +
                                   QuarterFromDate(MainFormAsm.iniSet.CbPriceLevelO);
            } else {
                rangeWork.Value2 = @"Составлена в ценах по состоянию на " +
                                   MainFormAsm.iniSet.CbPriceLevelL.Date.ToLongDateString();
            }
            RewriteFirstStringTable(sheet);
            // Подписи в конце страницы =======================================================
            if (MainFormAsm.iniSet.CbInsertSignOE) {
                // Всего по объектной смете
                Range findEnd = sheet.Cells.Find(@"Всего по объектной смете");
                int rowEnd = findEnd.Row + 1;
                Range www = sheet.Range["A" + rowEnd.ToString() + ":J" + ((int)(rowEnd + 15)).ToString()];
                www.UnMerge();
                www.Clear();
                var rowGip = rowEnd + 3;
                var rowBoss = rowEnd + 6;
                var rowMadeIn = rowEnd + 9;
                // вставим надписи и ФИО
                rangeWork = sheet.Cells[rowGip, "C"];
                rangeWork.Value2 = @"Главный инженер проекта :";
                rangeWork.HorizontalAlignment = XlHAlign.xlHAlignRight;
                rangeWork = sheet.Cells[rowBoss, "C"];
                rangeWork.Value2 = @"Руководитель группы смет :";
                rangeWork.HorizontalAlignment = XlHAlign.xlHAlignRight;
                rangeWork = sheet.Cells[rowMadeIn, "C"];
                rangeWork.Value2 = @"Составил :";
                rangeWork.HorizontalAlignment = XlHAlign.xlHAlignRight;

                rangeWork = sheet.Cells[rowGip, "D"];
                rangeWork.Value2 = "_____________________________" + MainFormAsm.iniSet.CbGipText;
                rangeWork.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                rangeWork = sheet.Cells[rowBoss, "D"];
                rangeWork.Value2 = "_____________________________" + MainFormAsm.iniSet.TbHeadDepartment;
                rangeWork.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                rangeWork = sheet.Cells[rowMadeIn, "D"];
                rangeWork.Value2 = "_____________________________" + MainFormAsm.iniSet.CbMadeInText;
                rangeWork.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                // Вставим картинки
                if (!MainFormAsm.iniSet.CbGip.Equals("")) {
                    InsertImage(ref sheet, rowGip, 5, MainFormAsm.iniSet.CbGipText);
                }
                if (!MainFormAsm.iniSet.TbHeadDepartment.Equals("")) {
                    InsertImage(ref sheet, rowBoss, 5, MainFormAsm.iniSet.TbHeadDepartment);
                }
                if (!MainFormAsm.iniSet.CbMadeIn.Equals("")) {
                    InsertImage(ref sheet, rowMadeIn, 5, MainFormAsm.iniSet.CbMadeInText);
                }
            }
            sheet.Name = @"ОС" + numberOE;
            return sheet.Name;
        }

        // Переделка начальных строк таблицы
        private void RewriteFirstStringTable(_Worksheet sheet) {
            Range col = sheet.Columns[2];
            col.ColumnWidth = "14";
            Range find = sheet.Cells.Find(@"Локальные сметные расчеты");
            if (find == null) {
                return;
            }
            int end = sheet.Cells.Find("Итого \"Локальные сметные расчеты\"").Row;
            int y = find.Row + 1;
            int x1 = find.Column + 1;
            int x2 = x1 + 1;
            for (int i = y; i < end; i++) {
                Range r1 = sheet.Cells[i, x1];
                Range r2 = sheet.Cells[i, x2];
                string sss = r2.Text;
                Regex pattern = new Regex(@"(?:\D*)(?<ss>((\d{2,4})[-|\.])*(\d{1,4}))");
                MatchCollection mc = pattern.Matches(sss);
                if (mc.Count == 0) {
                    continue;
                }

                string s1 = r2.Value2;
                string s2 = s1.Substring(0, s1.IndexOf(' '));
                s1 = s1.Substring(s1.IndexOf(' '));
                if (s1 != null && s2 != null) {
                    r1.Value2 = s2;
                    r2.Value2 = s1;
                }
            }
        }

        // Ресурсы. Обработка
        private string WorkWithExcelR(Worksheet sheet) {
            // Это наименование работ и т.д. ===============================================
            Range rangeWork = sheet.Range["A1"];
            if (MainFormAsm.iniSet.RbRes6) {
                rangeWork = sheet.Range[MainFormAsm.iniSet.tbRname6, MainFormAsm.iniSet.tbRname6];
            } else if (MainFormAsm.iniSet.RbRes7) {
                rangeWork = sheet.Range[MainFormAsm.iniSet.tbRname7, MainFormAsm.iniSet.tbRname7];
            }
            string nameWorks = RenameName(rangeWork.Value2);
            if (rangeWork.Value2 != null) {
                string sss = rangeWork.Value2.ToString();
                rangeWork.Value2 = RemoveBeginPos(sss);
            }
            // Это номер сметы ===============================================
            rangeWork = sheet.Cells.Find(@"ВЕДОМОСТЬ РЕСУРСОВ");
            if (rangeWork.Value2 != null) {
                string number = rangeWork.Value2;
                rangeWork.Value2 = number + " " + nameWorks;
            }
            // Это уровень цен ===============================================
            rangeWork = sheet.Cells.Find(@"по состоянию на");
            string price = rangeWork.Value2;
            rangeWork.Value2 = price + " " + QuarterFromDate(MainFormAsm.iniSet.CbPriceLevelL);
            //            rangeWork.Value2 = price + " " + cbPriceLevel.Text;
            rangeWork.HorizontalAlignment = HorizontalAlignment.Right;
            // Имя файла
            sheet.Name = @"Р" + nameWorks;
            // Это название стройки ===============================================
            rangeWork = sheet.Range["B2", "H2"];
            rangeWork.MergeCells = true;
            rangeWork.WrapText = true;
            rangeWork.Value2 = MainFormAsm.iniSet.TbNameBuilding;
            rangeWork.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            rangeWork.Font.Size = 12;
            rangeWork.NumberFormatLocal = "Основной";
            SetRowHeigths(ref sheet, ref rangeWork);
            return nameWorks;
        }

        private string WorkWithExcelVR(_Worksheet sheet) {
            // Это наименование работ и т.д. ===============================================
            var rangeWork = sheet.Range["C7", "C7"];
            string nameWorks = RenameName(rangeWork.Value2);
            // Убираем знак "№" из заголовка ===============================================
            rangeWork = sheet.Cells.Find(@"ВЕДОМОСТЬ ОБЪЕМОВ РАБОТ №");
            if (rangeWork.Value2 != null) {
                string number = rangeWork.Value2;
                rangeWork.Value2 = number.Remove(number.IndexOf("№", System.StringComparison.Ordinal));
            }
            // Уберем все лишнее сверху ===============================================
            var range5 = sheet.Range["A1", "E4"];
            range5.ClearContents();
            // Имя страницы
            sheet.Name = @"ВР" + nameWorks;
            return nameWorks;
        }

        private void SetRowHeigths(ref Worksheet ws, ref Range src) {
            Range test = ws.Cells[900, 100];
            int aa = src.EntireColumn.Count;
            double colWidth = 0;
            for (int i = 1; i <= aa; i++) {
                Range r = src.EntireColumn[i];
                colWidth = colWidth + r.ColumnWidth;
            }
            test.ColumnWidth = colWidth;
            test.Font.Size = 12;
            test.Value2 = src.Value2;
            test.WrapText = true;
            test.Rows.AutoFit();
            double h = test.RowHeight;
            h = Math.Ceiling(h / 10) * 10;
            src.RowHeight = h;
            test.Delete();
        }

    //    private float FloatTopPixelsCalculation(Range range) {
    //        Range r1 = range.Worksheet.Cells[range.Row + 1, range.Column];
    //        float floatTop1 = 0;
    //        for (var rNumber = 2; rNumber < r1.Row; rNumber++) {

    //            var cellHeight = Convert.ToSingle(r1.Worksheet.Cells[rNumber, r1.Column].RowHeight);
    //            floatTop1 = floatTop1 + cellHeight;
    //        }
    //        float floatTop = 0;
    //        for (var rNumber = 2; rNumber < range.Row; rNumber++) {
    //            var cellHeight = Convert.ToSingle(range.Worksheet.Cells[rNumber, range.Column].RowHeight);
    //            floatTop = floatTop + cellHeight;
    //        }
    //        return (floatTop + floatTop1) / 2;
    //    }

    //    private float FloatLeftPixelsCalculation(Range range) {
    //        float floatLeft = 0;
    //        for (var columnNumber = 1; columnNumber < range.Columns.Column; columnNumber++) {
    //            var cellWidth = Convert.ToSingle(range.Worksheet.Cells[range.Row, columnNumber].Width);
    //            floatLeft = floatLeft + cellWidth + 1;
    //        }
    //        return floatLeft;
    //    }
    }
}
