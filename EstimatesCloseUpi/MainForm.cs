using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace EstimatesName
{
    public partial class MainForm : Form {

        private List<Application> _list = null;
        private Ini _inifile = null;
        private const string ExtXLSX = @".xlsx";
        private string[] excessString = {
            "Сводный сметный расчет",
            "Объектная смета",
            "Полный локальный сметный расчёт",
            "Локальный сметный расчёт",
            "Локальный ресурсный сметный расчет"
        };
        private readonly ListViewColumnSorter _lvwColumnSorter;

        public MainForm() {
            InitializeComponent();
            _lvwColumnSorter = new ListViewColumnSorter();
            this.listViewProcess.ListViewItemSorter = _lvwColumnSorter;
        }

        private void FillListView() {
            //if (true) {
            //    _list = BovExcel.GetEnumRunningExcel(true);
            //} else {
            //    _list = BovExcel.GetEnumRunningExcel(false);
            //}
            _list = BovExcel.GetEnumRunningExcel(true);
            listViewProcess.Items.Clear();
            foreach (var ex in _list) {
                listViewProcess.Items.Add(new ListViewItem(new[] { ex.Hwnd.ToString(), ex.ActiveWorkbook.Name }));
            }
            lblNumTable.Text = String.Format("Кол-во открытых таблиц: {0}", listViewProcess.Items.Count);
        }

        private void button2_Click(object sender, EventArgs e) {
            Close();
        }

        private void MainForm_Shown(object sender, EventArgs e) {
            FillListView();
        }

        private void MainForm_Load(object sender, EventArgs e) {
            _inifile = new Ini(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\estimate.ini");
            tbWorkPath.Text = _inifile.IniReadValue(@"Global", @"workPath");
            _lvwColumnSorter.SortColumn = 1;
            _lvwColumnSorter.Order = SortOrder.Ascending;
            listViewProcess.Sort();
            lblNumTable.Text = String.Format("Кол-во открытых таблиц: {0}", listViewProcess.Items.Count);
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e) {
            _inifile.IniWriteValue(@"Global", @"workPath", tbWorkPath.Text);
        }

        // Сортировка колонок
        private void listView1_ColumnClick(object sender, ColumnClickEventArgs e) {
            // Determine if clicked column is already the column that is being sorted.
            if (e.Column == _lvwColumnSorter.SortColumn) {
                // Reverse the current sort direction for this column.
                if (_lvwColumnSorter.Order == SortOrder.Ascending) {
                    _lvwColumnSorter.Order = SortOrder.Descending;
                } else {
                    _lvwColumnSorter.Order = SortOrder.Ascending;
                }
            } else {
                // Set the column number that is to be sorted; default to ascending.
                _lvwColumnSorter.SortColumn = e.Column;
                _lvwColumnSorter.Order = SortOrder.Ascending;
            }

            // Perform the sort with these new sort options.
            this.listViewProcess.Sort();
        }

        private void показатьToolStripMenuItem_Click(object sender, EventArgs e) {
            ListView.SelectedListViewItemCollection selectedLvi = this.listViewProcess.SelectedItems;
            if (selectedLvi.Count > 1) {
                MessageBox.Show(@"Выбрано больше одного документа!", @"Внимание!", MessageBoxButtons.OK);
            } else {
                string text = selectedLvi[0].SubItems[0].Text;
                foreach (var application in _list) {
                    if (application.Hwnd.ToString(CultureInfo.InvariantCulture).Equals(text)) {
                        application.Visible = true;
                        break;
                    }
                }
            }
            FillListView();
        }

        private void закрытьToolStripMenuItem_Click(object sender, EventArgs e) {
            var selectedLvi = this.listViewProcess.SelectedItems;
            if (selectedLvi.Count > 1) {
                MessageBox.Show(@"Выбрано больше одного документа!", @"Внимание!", MessageBoxButtons.OK);
            } else {
                var text = selectedLvi[0].SubItems[0].Text;
                foreach (var application in _list.Where(application => application.Hwnd.ToString(CultureInfo.InvariantCulture).Equals(text))) {
                    BovExcel.FinishExcel(application);
                    break;
                }
            }
            FillListView();
        }

        private void закрытьССохранениемToolStripMenuItem_Click(object sender, EventArgs e) {
            var selectedLvi = this.listViewProcess.SelectedItems;
            if (selectedLvi.Count > 1) {
                MessageBox.Show(@"Выбрано больше одного документа!", @"Внимание!", MessageBoxButtons.OK);
            } else if (selectedLvi.Count == 1) {
                var text = selectedLvi[0].SubItems[0].Text;
                foreach (var application in _list.Where(application => application.Hwnd.ToString(CultureInfo.InvariantCulture).Equals(text))) {
                    FillWorksheetAndSave(application);
                    BovExcel.FinishExcel(application);
                    break;
                }
            } else if (selectedLvi.Count == 0)  {
                MessageBox.Show(@"Не выбрано ни одного документа!", @"Внимание!", MessageBoxButtons.OK);
            }
            FillListView();
        }

        private void закрытьВсеToolStripMenuItem_Click(object sender, EventArgs e) {
            foreach (var application in _list) {
                BovExcel.FinishExcel(application);
                FillListView();
            }
        }

        private void закрытьВСЕССохранениемToolStripMenuItem_Click(object sender, EventArgs e) {
            foreach (var application in _list) {
                FillWorksheetAndSave(application);
                BovExcel.FinishExcel(application);
                FillListView();
            }
        }

        private void FillWorksheetAndSave(Application application) {
            application.DisplayAlerts = false;
            // Обработка имен окон -> имя файла
            string nameWorkbook = application.ActiveWorkbook.Name.Replace(".", "").Replace("_", " ");
            string fullpath = Path.GetFileNameWithoutExtension(nameWorkbook);
            foreach (String s in excessString) {
                int a = fullpath.IndexOf(s);
                if (a > 3) {
                    fullpath = fullpath.Remove(a - 3);
                    break;
                }
            }
            //
            fullpath = Path.Combine(tbWorkPath.Text, fullpath + ExtXLSX);
            application.ActiveWorkbook.SaveAs(fullpath, XlFileFormat.xlOpenXMLWorkbook);
        }

        private void button3_Click(object sender, EventArgs e) {
            dlgPath.SelectedPath = tbWorkPath.Text;
            dlgPath.ShowDialog();
            tbWorkPath.Text = dlgPath.SelectedPath;
        }

        private void tbWorkPath_TextChanged(object sender, EventArgs e) {
            _inifile.IniWriteValue("Global", "workPath", tbWorkPath.Text);
        }

        private void обновитьСписокToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FillListView();
            if (_list.Count == 0)
            {
                MessageBox.Show(@"Ни одного экземпляра EXCEL на запущено.", @"Внимание!", MessageBoxButtons.OK);
            }
            lblNumTable.Text = String.Format("Кол-во открытых таблиц: {0}", listViewProcess.Items.Count);
        }
    }

}
