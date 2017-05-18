using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Math;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.Win32;
using Application = Microsoft.Office.Interop.Excel.Application;
using Page = DocumentFormat.OpenXml.Spreadsheet.Page;
using Path = System.IO.Path;
using Shape = Microsoft.Office.Interop.Excel.Shape;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;
using XlHAlign = Microsoft.Office.Interop.Excel.XlHAlign;
using PrinterSettings = System.Drawing.Printing.PrinterSettings;

namespace EstimatesAssembly {
    class BookEstimates {
        private const int PixelW = 100;
        private const int PixelH = 50;
        private string propertyDocName = "Comments";
        private string propertyDocValue = "собранная книга со сметами";
        private ProgressBar _pgBar;

        public ProgressBar PgBar {
            get { return _pgBar; }
            set { _pgBar = value; }
        }

        private string _nameBook;
        private string _pathBook;
        public Application Ex;
        public Workbook Wb;
        public Workbook TmpWb;

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
                    var prn = String.Format("{0} ({1})", printer, port);
                    Ex.ActivePrinter = prn;
                }
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
            _pgBar.Maximum = selectedItems.Length;
            foreach (string selectedItem in selectedItems) {
                TmpWb = Ex.Workbooks.Open(selectedItem);
                foreach (Worksheet sheet in TmpWb.Sheets) {
                    _pgBar.Value += 1;
                    switch (FindTypeSheet(sheet)) {
                        case 1:
                            WorkWithExcelLs(sheet); // Локальная смета
                            break;
                        case 2:
                            WorkWithExcelOs(sheet); // Объектная смета
                            break;
                        case 3:
                            WorkWithExcelR(sheet); // Ресурсная смета
                            break;
                        case 4:
                            WorkWithExcelVR(sheet); // Ведомость ресурсов
                            break;
                        case 5:
                            WorkWithExcelSSR(sheet); // Сводный сметный расчет
                            break;
                        case 6:
                            WorkWithExcelLRS(sheet); // Локальная ресурсная смета
                            break;
                    }
                    sheet.Copy(Type.Missing, Wb.ActiveSheet);
                }
                TmpWb.Close();
            }
            _pgBar.Value = 0;

            foreach (string myvar in GetListSheet()) {
                if (myvar.Contains("Лист")) {
                    Wb.Sheets[myvar].Delete();
                }
            }
            SetActivePrinterPDF();
            SetDocumentProperty(propertyDocName, propertyDocValue);
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
            _pgBar.Maximum = selectedItems.Count;
            foreach (ListViewItem selectedItem in selectedItems) {
                _pgBar.Value += 1;
                Worksheet worksheet = Wb.Sheets[selectedItem.Text];
                if (worksheet.Visible == XlSheetVisibility.xlSheetHidden) {
                    worksheet.Visible = XlSheetVisibility.xlSheetVisible;
                }
                if (Wb.Sheets.Count == 1) {
                    Wb.Sheets.Add();
                }
                Wb.Sheets[selectedItem.Text].Delete();
            }
            _pgBar.Value = 0;
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
            foreach ( Worksheet sheet in workbook.Sheets) {
                list.Add(sheet.Name);
            }
            return list;
        }

        // Сохранение тома
        public void SaveWorkbook() {
            string fullname = Path.Combine(_pathBook, _nameBook + @".xlsx");
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
                    XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
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
                MessageBox.Show(@"Не выбрано ни одной сметы!", @"Внимание!");
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
            _pgBar.Maximum = Ex.ActiveWorkbook.Sheets.Count;
            foreach (Worksheet ws in Ex.ActiveWorkbook.Sheets) {
                _pgBar.Value += 1;
                list.Add(ws.Name);
            }
            _pgBar.Value = 0;
            list.Sort(Compare);
            Workbook wb = Ex.ActiveWorkbook;
            _pgBar.Maximum = list.Count + 1;
            foreach (string str in list) {
                _pgBar.Value += 1;
                Worksheet ws = wb.Sheets[str];
                ws.Move(Wb.Sheets[list.IndexOf(str) + 1], Type.Missing);
            }
            _pgBar.Value = 0;
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

        // Дополнительная обработка таблиц. Смещение последнего разрыва
        public void AdaptionSheets() {
            Range range;
            Workbook mainBook = Ex.ActiveWorkbook;
            if (mainBook == null) {
                return;
            }
            _pgBar.Maximum = mainBook.Sheets.Count;
            foreach (Worksheet worksheet in mainBook.Sheets) {
                if (!worksheet.Name.Equals(@"Оглавление")) {
                    _pgBar.Value += 1;
                    worksheet.Activate();
                    Ex.ActiveWindow.View = XlWindowView.xlPageLayoutView;
                    HPageBreaks hbreaks = worksheet.HPageBreaks;
                    int pageCount = hbreaks.Count;
                    if (pageCount != 0) {
                        range = hbreaks.Item[pageCount].Location;
                        int t = FindLastRow(worksheet);
                        int t1 = range.Row;
                        int diff = t - t1;
                        Ex.ActiveWindow.View = XlWindowView.xlPageBreakPreview;
                        if (diff < 12 && diff > 0) {
                            HPageBreak brr = hbreaks.Item[pageCount];
                            brr.Location = worksheet.Range["A" + Convert.ToString(t1 - 12)];
                        }
                    }
                    Release(hbreaks);
                }
            }
            _pgBar.Value = 0;
        }

        // Найти последний используемый столбец
        private int FindRightColumn(Worksheet worksheet) {
            Range range = worksheet.UsedRange;
            return range.Columns.Count;
        }

        // Найти последнюю используемую строку
        private int FindLastRow(Worksheet worksheet) {
            int lastCol = worksheet.UsedRange.Columns.Count;
            int fullRow = worksheet.Rows.Count;
            int lastRow = worksheet.Cells[fullRow, 1].End(XlDirection.xlUp).Row;

            for (int i = 0; i < lastRow; i++) {
                Range range = worksheet.Cells[i + 1, 1];
                var mergeCells = range.MergeCells;
                if (mergeCells && (range.MergeArea.Columns.Count == lastCol)) {
                    lastRow++;
                }
            }
            return lastRow;
        }

        // перенумерация страниц
        public void NumberingPage() {
            Workbook mainBook = Ex.ActiveWorkbook;
            if (mainBook == null) {
                return;
            }
            foreach (Worksheet worksheet in mainBook.Sheets) {
                if (worksheet.Name.Contains(@"Лист")) worksheet.Delete();
            }
            // Включим разрывы страниц
            _pgBar.Maximum = mainBook.Sheets.Count;
            foreach (Worksheet worksheet in mainBook.Sheets) {
                _pgBar.Value += 1;
                worksheet.PageSetup.Orientation = XlPageOrientation.xlLandscape;
                worksheet.PageSetup.Zoom = false;
                worksheet.PageSetup.FitToPagesWide = 1;
                worksheet.PageSetup.FitToPagesTall = 999;
                worksheet.PageSetup.TopMargin = Ex.CentimetersToPoints(0.5);
                worksheet.PageSetup.BottomMargin = Ex.CentimetersToPoints(1.4);
                worksheet.PageSetup.HeaderMargin = 0.0;
                worksheet.PageSetup.RightFooter = "&P";
                worksheet.PageSetup.LeftHeader = "";
                worksheet.PageSetup.CenterHeader = "";
                worksheet.PageSetup.RightHeader = "";
                worksheet.Columns[5].ColumnWidth = 8f;
                worksheet.PageSetup.PrintTitleRows = "";
                Ex.ActiveWindow.View = XlWindowView.xlPageBreakPreview;
            }
            _pgBar.Value = 0;

            int x = 1;
            _pgBar.Maximum = mainBook.Sheets.Count;
            foreach (Worksheet worksheet in mainBook.Sheets) {
                _pgBar.Value += 1;
                worksheet.Select();
                Ex.ActiveWindow.View = XlWindowView.xlPageBreakPreview;
                worksheet.PageSetup.RightFooter = "&P";
                worksheet.PageSetup.LeftHeader = " ";
                worksheet.PageSetup.CenterHeader = " ";
                worksheet.PageSetup.RightHeader = " ";
                worksheet.PageSetup.FirstPageNumber = x;
                int i = worksheet.PageSetup.FirstPageNumber;
                int j = worksheet.PageSetup.Pages.Count;
                x = i + j;
            }
            _pgBar.Value = 0;
            AdaptionSheets();
            
        }

        // Вставить картинку
        private void InsertImage(ref Worksheet sheet, int y, int x, string fio) {
            char[] charsToTrim = { '\n', '\r', ' ' };
            Shape shape = null;
            Range range = sheet.Cells[y, x];
            fio = fio.TrimEnd(charsToTrim);
            float xx = (float)((double)range.Left - 10);
            float yy = (float)((double)range.Top - 20);
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

        // Преобразовать Ф И.О. к виду Ф_И_О
        private static string ConvertName(string name) {
            string n = name.Replace(".", "_").Replace(" ", "_");
            n = n.Substring(0, n.Length - 1);
            return n;
        }

        // Получаем (если нужно) квартал из даты
        private string QuarterFromDate(DateTime value) {
            int a = DateAndTime.DatePart(DateInterval.Quarter, value);
            int b = DateAndTime.DatePart(DateInterval.Year, value);
            if (MainFormAsm.iniSet.CbQuarter) {
                return String.Format("{0}-й квартал {1} года.", a, b);
            } else {
                return value.ToString("dd MMMM yyyy", CultureInfo.CreateSpecificCulture("ru-RU"));
            }
        }


        public static string ConvertRSName(string name) {
            Regex pattern = new Regex(@"(?:\D*)(?<ss>((\d{2,4})[-|\.])*(\d{1,4}))");
            MatchCollection mc = pattern.Matches(name);
            if (mc.Count > 0) {
                GroupCollection groups = mc[0].Groups;
                name = name.Remove(0, groups["ss"].Value.Length);
                return name;
            } else {
                return name;
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

        // Доп. обработка объектных смет. 
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
                price.Value2 = @"Составлен(а) в ценах по состоянию на " + QuarterFromDate(MainFormAsm.iniSet.CbPriceLevelO);
                sheet.Name = @"СС00";
            } else {
                Range price = sheet.Range["B18"];
                price.Value2 = @"Составлен(а) в ценах по состоянию на 01.01.2000";
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
                    if (!MainFormAsm.iniSet.TbGipMain.Equals("")) {
                        InsertImage(ref sheet, rowGip, 5, MainFormAsm.iniSet.CbGipText);
                    }
                    if (!MainFormAsm.iniSet.TbHeadDepartment.Equals("")) {
                        InsertImage(ref sheet, rowBoss, 5, MainFormAsm.iniSet.TbHeadDepartment);
                    }
                    if (!MainFormAsm.iniSet.TbBuilderMain.Equals("")) {
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

        // Локальные сметы. Обработка
        private string WorkWithExcelLs(Worksheet sheet) {
            sheet.UsedRange.Font.Name = "Times New Roman";
            sheet.Range["A1:Q5"].Clear();
            string numberEstimate = sheet.Range["G9"].Text;
            string nameObject = MainFormAsm.iniSet.TbNameBuilding;
            numberEstimate = numberEstimate.Substring(numberEstimate.LastIndexOf("№") + 2);
            string tmpNameObject = sheet.Range["G6"].Text;
            Range tmprange = sheet.Range["A6:Q6"];
            tmprange.Merge();
            tmprange.Font.Bold = true;
            tmprange.Font.Underline = true;
            tmprange.EntireRow.RowHeight = 20;
            tmprange.EntireRow.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            tmprange.Value2 = tmpNameObject;

            tmprange = sheet.Range["A7:Q7"];
            tmprange.Merge();
            tmprange.Font.Italic = true;
            tmprange.Merge();
            tmprange.Value2 = "наименование стройки";

            tmprange = sheet.Range["G9"];
            tmprange.Font.Name = "Times New Roman";
            tmprange.Font.Bold = true;

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
            return numberEstimate;
        }

        // Объектные сметы. Обработка
        private string WorkWithExcelOs(Worksheet sheet) {
            string tmp;
            sheet.UsedRange.Font.Name = "Times New Roman";
            // Затрем все лишнее сверху ===============================================
            ((Range)sheet.Range["J1"]).Clear();
            string tmpNameObj = sheet.Range["E2"].Text;
            sheet.Range["A2:J2"].Merge();
            sheet.Range["A2:J2"].Font.Bold = true;
            sheet.Range["A2:J2"].Font.Underline = true;
            sheet.Range["A2:J2"].Value2 = tmpNameObj;

            tmp = sheet.Range["E3"].Text;
            sheet.Range["A3:J3"].Merge();
            sheet.Range["A3"].Value2 = tmp;
            sheet.Range["A3"].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            sheet.Range["A3"].Font.Italic = true;
            tmp = sheet.Range["E6"].Text;
            sheet.Range["A6:J6"].Merge();
            sheet.Range["A6"].Value2 = tmp;
            sheet.Range["A6"].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            sheet.Range["A6"].Font.Italic = true;
            tmp = sheet.Range["E9"].Text;
            sheet.Range["A9:J9"].Merge();
            sheet.Range["A9"].Value2 = tmp;
            sheet.Range["A9"].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            sheet.Range["A9"].Font.Italic = true;

            string numberOE = ((Range)sheet.Range["G5"]).Text;
            sheet.Range["G5"].Clear();
            sheet.Range["E5"].Value2 += numberOE;
            sheet.Range["A5:J5"].Merge();
            sheet.Range["A5:J5"].Font.Underline = true;
            sheet.Range["A5:J5"].Font.Bold = true;

            tmp = sheet.Range["D8"].Text;
            Regex pattern = new Regex(@"(?:\D*)(?<ss>((\d{2,4})[-|\.])*(\d{1,4}))");
            MatchCollection mc = pattern.Matches(tmp);
            if (mc.Count == 1) {
                tmp = tmp.Substring(mc[0].Value.Length).Trim();
            }
            sheet.Range["A8:J8"].Clear();
            sheet.Range["A8:J8"].Merge();
            sheet.Range["A8"].Value2 = tmp;
            sheet.Range["A8"].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            sheet.Range["A8"].Font.Name = "Times New Roman";
            sheet.Range["A8"].Font.Bold = true;
            sheet.Range["A8"].Font.Underline = true;

            Range rangeWork;
            string nameWorks = numberOE;
            // Это уровень цен ===============================================
            string stmp = @"Составлен(а) в ценах по состоянию на ";
            rangeWork = sheet.Range["C14"];
            if (MainFormAsm.iniSet.CbQuarter) {
                rangeWork.Value2 = stmp + QuarterFromDate(MainFormAsm.iniSet.CbPriceLevelO);
            } else {
                rangeWork.Value2 = stmp + MainFormAsm.iniSet.CbPriceLevelL.Date.ToLongDateString();
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

                rangeWork = sheet.Cells[rowGip, "D"];
                rangeWork.Value2 = "_____________________________" + MainFormAsm.iniSet.CbGipText;
                rangeWork.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                rangeWork = sheet.Cells[rowBoss, "D"];
                rangeWork.Value2 = "_____________________________" + MainFormAsm.iniSet.TbHeadDepartment;
                rangeWork.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                if (!MainFormAsm.iniSet.TbGipMain.Equals("")) {
                    InsertImage(ref sheet, rowGip, 5, MainFormAsm.iniSet.CbGipText);
                }
                if (!MainFormAsm.iniSet.TbHeadDepartment.Equals("")) {
                    InsertImage(ref sheet, rowBoss, 5, MainFormAsm.iniSet.TbHeadDepartment);
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
            rangeWork = sheet.Range["C8", "C8"];
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

        // Ведомость объемов работ
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

        private void Release(object sender) {
            try {
                if (sender != null) {
                    Marshal.ReleaseComObject(sender);
                    sender = null;
                }
            } catch (Exception) {
                sender = null;
            }
        }

        private object GetDocumentProperty(string propertyName) {
            object returnVal = null;
            dynamic properties = Ex.ActiveWorkbook.BuiltinDocumentProperties;
            foreach (dynamic property in properties) {
                string name = property.Name;
                if (name.Equals(propertyName)) {
                    returnVal = property.Value;
                }
            }
            string test = returnVal.ToString();
            return returnVal;
        }

        protected void SetDocumentProperty(string propertyName, string propertyValue) {
            bool propertyExists = false;
            dynamic properties = Ex.ActiveWorkbook.BuiltinDocumentProperties;
            foreach (dynamic prop in properties) {
                if (prop.Name == propertyName) {
                    prop.Value = propertyValue;
                    propertyExists = true;
                    break;
                }
            }
            if (!propertyExists) {
                properties.Add(propertyName, false, MsoDocProperties.msoPropertyTypeString, propertyValue, Type.Missing);
            }
        }

    }
}
