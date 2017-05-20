using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using System.Xml.Serialization;
using EstimatesName;
using EstimatesAssembly;
using Microsoft.VisualBasic;
using VBasic = Microsoft.VisualBasic.Interaction;

namespace EstimatesAssembly {
    public partial class MainFormAsm : Form {

        private BookEstimates _book;
        private VolumeAsm _volumeAsm;
        private readonly string _configfile;
        private initBookmark InitBookmark = new initBookmark();
        private readonly ListViewColumnSorter _lvwColumnSorter = new ListViewColumnSorter();
        public static Config iniSet = new Config();
        public Dictionary<string, string> mapBookmak = new Dictionary<string, string>();

        Object missingObj = System.Reflection.Missing.Value;

        public MainFormAsm() {
            InitializeComponent();
            _book = new BookEstimates();
            _volumeAsm = new VolumeAsm();
            _configfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\estimate.xml";
            _book.PgBar = pgBar;
            this.lstSheet.ListViewItemSorter = _lvwColumnSorter;
            // перехватчики глобальных исключений
            AppDomain.CurrentDomain.UnhandledException += delegate (object sender, UnhandledExceptionEventArgs args)
            {
                allQuit();
                Trace.WriteLine("Global exception: " + args.ExceptionObject.ToString());
            };
            Application.ThreadException += delegate (Object sender, ThreadExceptionEventArgs args)
            {
                allQuit();
                Trace.WriteLine("Global exception: " + args.Exception.ToString());
                Environment.Exit(0);
            };

        }

        public void allQuit() {
            if (_book != null) {
                if (_book.Ex != null) {
                    _book.Ex.Application.Quit();
                }
            }
            if (_volumeAsm != null) {
                if (_volumeAsm._appExcel != null) {
                    _volumeAsm._appExcel.Quit();
                }
                if (_volumeAsm._appWord != null) {
                    _volumeAsm._appWord.Quit();
                }
            }
        }

        // Задать папку для вывода результата сборки
        private void btnEstimatePathAndName_Click(object sender, EventArgs e) {
            folderBrowserDialog.SelectedPath = txtEsimatePath.Text;
            folderBrowserDialog.ShowDialog();
            txtEsimatePath.Text = folderBrowserDialog.SelectedPath;
            SaveConfig();
        }

        // Задать путь с изображениями
        private void btnImagePath_Click(object sender, EventArgs e) {
            folderBrowserDialog.SelectedPath = txtImagePath.Text;
            folderBrowserDialog.ShowDialog();
            txtImagePath.Text = folderBrowserDialog.SelectedPath;
            SaveConfig();
        }

        // Выход из программы
        private void button1_Click(object sender, EventArgs e) {
            _book.CloseBook();
            SaveConfig();
            Close();
        }

        // Загрузка главной формы
        private void MainFormAsm_Load(object sender, EventArgs e) {
            ReadConfig(); // читаем настройки
            mapBookmak = InitBookmark.fillBookmark();
            ChangeNameBook(); // задаем название книги
            ListRefresh();
            // создаем временный шаблон оглавления
        }

        // Закрытие главной формы
        private void MainFormAsm_FormClosing(object sender, FormClosingEventArgs e) {
            SaveConfig(); // сохраняем настройки
        }

        // Сохранение настроек в классе и сериализация в XML
        private void SaveConfig() {
            iniSet.TxtEsimatePath = txtEsimatePath.Text;
            iniSet.TxtImagePath = txtImagePath.Text;
            iniSet.TbNameBuilding = tbNameBuilding.Text;
            iniSet.TbNameObject = tbNameObject.Text;
            iniSet.TbCodeObject = tbCodeObject.Text;
            iniSet.NumVolumeNumber = numVolumeNumber.Text;
            iniSet.TbInventoryNumber = tbInventoryNumber.Text;
            iniSet.CbStageDevelope = cbStageDevelope.Text;
            iniSet.TbGipMain = tbGipMain.Text;
            iniSet.TbBuilderMain = tbBuilderMain.Text;
            iniSet.TbChiefEngineer = tbChiefEngineerMain.Text;
            iniSet.TbHeadDepartment = tbHeadDepartmentMain.Text;
            iniSet.DateToStamp = dateToStamp.Value;
            iniSet.CbPriceLevelL = cbPriceLevelL.Value;
            iniSet.CbPriceLevelO = cbPriceLevelO.Value;
            iniSet.CbQuarter = cbQuarter.Checked;
            iniSet.CbInsertSignOE = cbInsertSignOE.Checked;
            iniSet.CbInsertSignLE = cbInsertSignLE.Checked;
            iniSet.CbInsertSignSS = cbInsertSignSS.Checked;
            iniSet.CbInsertSignLR = cbInsertSignLR.Checked;
            iniSet.CbSort = chbSort.Checked;
            iniSet.CbNumeric = chbNumeric.Checked;
            iniSet.CbSave = chbSave.Checked;
            iniSet.TbCustomer = tbCustomer.Text;
            iniSet.TbCertificate = tbCertificate.Text;
            iniSet.TbYearTitle = tbYearTitul.Text;
            iniSet.TbWorkFolder = tbWorkFolder.Text;

            iniSet.TbChiefPsition = tbChiefPosition.Text;
            iniSet.TbChiefFio = tbChiefFio.Text;
            iniSet.TbGipFio = tbGipFio.Text;
            iniSet.TbSectionNumber = tbSectionNumber.Text;
            iniSet.TbVolCount = tbVolCount.Text;
            iniSet.NumberFirstPage = numberFirstPage.Text;
            using (Stream writer = new FileStream(_configfile, FileMode.Create)) {
                var serializer = new XmlSerializer(typeof(Config));
                serializer.Serialize(writer, iniSet);
            }
        }

        // Восстановление настроек из файла в класс (десериализация)
        private void ReadConfig() {
            if (File.Exists(_configfile)) {
                using (Stream stream = new FileStream(_configfile, FileMode.Open)) {
                    var serializer = new XmlSerializer(typeof(Config));
                    iniSet = (Config)serializer.Deserialize(stream);
                    txtEsimatePath.Text = iniSet.TxtEsimatePath;
                    txtImagePath.Text = iniSet.TxtImagePath;
                    tbNameBuilding.Text = iniSet.TbNameBuilding;
                    tbNameObject.Text = iniSet.TbNameObject;
                    tbCodeObject.Text = iniSet.TbCodeObject;
                    numVolumeNumber.Text = iniSet.NumVolumeNumber;
                    tbInventoryNumber.Text = iniSet.TbInventoryNumber;
                    cbStageDevelope.Text = iniSet.CbStageDevelope;
                    tbChiefEngineerMain.Text = iniSet.TbChiefEngineer;
                    tbHeadDepartmentMain.Text = iniSet.TbHeadDepartment;
                    dateToStamp.Value = iniSet.DateToStamp;
                    cbPriceLevelL.Value = iniSet.CbPriceLevelL;
                    cbPriceLevelO.Value = iniSet.CbPriceLevelO;
                    cbQuarter.Checked = iniSet.CbQuarter;
                    cbInsertSignOE.Checked = iniSet.CbInsertSignOE;
                    cbInsertSignLE.Checked = iniSet.CbInsertSignLE;
                    cbInsertSignSS.Checked = iniSet.CbInsertSignSS;
                    cbInsertSignLR.Checked = iniSet.CbInsertSignLR;
                    chbSort.Checked = iniSet.CbSort;
                    chbNumeric.Checked = iniSet.CbNumeric;
                    chbSave.Checked = iniSet.CbSave;
                    tbCustomer.Text = iniSet.TbCustomer;
                    tbCertificate.Text = iniSet.TbCertificate;
                    tbYearTitul.Text = iniSet.TbYearTitle;
                    tbWorkFolder.Text = iniSet.TbWorkFolder;
                    tbChiefPosition.Text = iniSet.TbChiefPsition;
                    tbChiefFio.Text = iniSet.TbChiefFio;
                    tbGipFio.Text = iniSet.TbGipFio;
                    tbSectionNumber.Text = iniSet.TbSectionNumber;
                    tbVolCount.Text = iniSet.TbVolCount;
                    numberFirstPage.Text = iniSet.NumberFirstPage;
                }
            }
        }

        // Перерисовка таба
        private void tabPageEstimate_Paint(object sender, PaintEventArgs e) {
            ChangeNameBook();
        }

        // Изменение наименование книги
        private void ChangeNameBook() {
            _book.NameBook = @"Том-" + numVolumeNumber.Text + "-" + numVolumeNumber.Text;
            _book.PathBook = txtEsimatePath.Text + "\\";
        }

        // Добавить в список файл со сметой или еще с чем нибудь
        private void btnAddSheet_Click(object sender, EventArgs e) {
            dlgOpenFile.Filter = @"Excel files|*.xlsx;*.xls";
            dlgOpenFile.ShowDialog();
            if (dlgOpenFile.FileNames.Equals("")) {
                return;
            }
            _book.AddSheetNew(dlgOpenFile.FileNames);
            if (chbSort.Checked) {
                _book.SortWorksheets();
            }
            if (chbNumeric.Checked) {
                _book.NumberingPage();
            }
            if (chbSave.Checked) {
                _book.SaveWorkbook();
            }
            FillListSheet(_book.GetListSheet());
        }

        // Заполнение списка смет
        private void FillListSheet(IEnumerable<string> list) {
            lstSheet.Items.Clear();
            if (list == null) return;
            foreach (var str in list) {
                lstSheet.Items.Add(new ListViewItem(new[] { str }));
            }
        }

        // Сохранение книги
        private void button2_Click(object sender, EventArgs e) {
            _book.SaveWorkbook();
        }

        // Удаление из списка одного или нескольких элементов
        private void btnDelSheet_Click(object sender, EventArgs e) {
            if (lstSheet.SelectedItems.Count == 0) {
                MessageBox.Show(@"Пустая книга.");
                return;
            }
            _book.DeleteSheet(lstSheet.SelectedItems);
            FillListSheet(_book.GetListSheet());
        }

        // Показать том(сборку)
        private void checkBox1_CheckedChanged(object sender, EventArgs e) {
            if (cbShowExcel.Checked) {
                _book.ShowExcel(true);
            } else {
                _book.ShowExcel(false);
            }
        }

        // Переместить элемент в таблице вверх на одну позицию
        private void btnUpSheet_Click(object sender, EventArgs e) {
            if (lstSheet.SelectedItems.Count > 1) {
                MessageBox.Show(@"Выбрано болше одной сметы");
                return;
            } else if (lstSheet.SelectedItems.Count == 0) {
                MessageBox.Show(@"Пустая книга.");
                return;
            }
            string si = lstSheet.SelectedItems[0].Text;
            _book.MoveUpsheet(lstSheet.SelectedItems);
            FillListSheet(_book.GetListSheet());
            lstSheet.Focus();
            lstSheet.FindItemWithText(si).Selected = true;
        }

        // Переместить элемент в таблице вниз на одну позицию
        private void btnDownSheet_Click(object sender, EventArgs e) {
            if (lstSheet.SelectedItems.Count > 1) {
                MessageBox.Show(@"Выбрано болше одной сметы");
                return;
            } else if (lstSheet.SelectedItems.Count == 0) {
                MessageBox.Show(@"Пустая книга.");
                return;
            }
            string si = lstSheet.SelectedItems[0].Text;
            _book.MoveDownsheet(lstSheet.SelectedItems);
            FillListSheet(_book.GetListSheet());
            lstSheet.Focus();
            lstSheet.FindItemWithText(si).Selected = true;
        }

        // Пересортировка элементов в таблице
        private void btnSortSheet_Click(object sender, EventArgs e) {
            // сортировка таблиц в книге...
            _book.SortWorksheets();
            FillListSheet(_book.GetListSheet());
        }

        // Перечитать список элементов
        private void button1_Click_1(object sender, EventArgs e) {
            FillListSheet(_book.GetListSheet());
            lstSheet.Focus();
        }

        // Перенумерация сборки
        private void btnNumbering_Click(object sender, EventArgs e) {
            _book.NumberingPage();
            FillListSheet(_book.GetListSheet());
        }

        private void button2_Click_1(object sender, EventArgs e) {
            SaveConfig();
        }

        private void MainFormAsm_KeyDown(object sender, KeyEventArgs e) {
            if (e.KeyData == (Keys.Insert)) {
                btnAddSheet_Click(sender, e);
            }
        }

        private void btnSetWorkFolder_Click(object sender, EventArgs e) {
            folderBrowserDialog.SelectedPath = tbWorkFolder.Text;
            folderBrowserDialog.ShowDialog();
            tbWorkFolder.Text = folderBrowserDialog.SelectedPath;
            SaveConfig();
            ListRefresh();
        }

        private void listViewWithReordering_MouseDoubleClick(object sender, MouseEventArgs e) {
            string file = tbWorkFolder.Text + @"\"+ listViewWithReordering.SelectedItems[0].Text;
            Process.Start(file);
        }

        private void btnAddFile_Click(object sender, EventArgs e) {
            dlgOpenFile.Filter = @"Файлы для книги|*.xlsx;*.xls;*.docx;*.doc;*.pdf"; 
            dlgOpenFile.ShowDialog();
            string[] files = dlgOpenFile.FileNames;
            if (files.Length != 0) {
                foreach (var file in files) {
                    FileInfo fileInfo = new FileInfo(file);
                    string source = file;
                    string dest = tbWorkFolder.Text + @"\" + fileInfo.Name;
                    File.Copy(source, dest, true);
                }
                ListRefresh();
            }
        }

        private void btnDeleteFile_Click(object sender, EventArgs e) {
            DialogResult dialogResult;
            if (listViewWithReordering.SelectedItems.Count == 0) {
                MessageBox.Show(@"Ни одного файла не выбрано!", @"Внимание!");
                return;
            }
            if (listViewWithReordering.SelectedItems.Count > 1) {
                dialogResult = MessageBox.Show(@"Вы действительно хотите удалить эти файлы?", @"Внимание!", MessageBoxButtons.YesNo);
            } else {
                dialogResult = MessageBox.Show(@"Вы действительно хотите удалить этот файл?", @"Внимание!", MessageBoxButtons.YesNo);
            }
            if (dialogResult == DialogResult.No) return;
            foreach (ListViewItem selectedItem in listViewWithReordering.SelectedItems) {
                string file = tbWorkFolder.Text + @"\" + selectedItem.Text;
                if (File.Exists(file)) File.Delete(file);
            }
            ListRefresh();
        }

        private void btnRefresh_Click(object sender, EventArgs e) {
            ListRefresh();
        }

        private void btnRename_Click(object sender, EventArgs e) {
            if (listViewWithReordering.SelectedItems.Count == 0) {
                MessageBox.Show(@"Ни одного файла не выбрано!", @"Внимание!");
                return;
            } else if (listViewWithReordering.SelectedItems.Count > 1) {
                MessageBox.Show(@"Выбрано больше одного файла! Так нельзя!", @"Внимание!");
                return;
            }
            string oldName = listViewWithReordering.SelectedItems[0].Text;
            FileInfo fileInfo = new FileInfo(tbWorkFolder.Text + @"\" + oldName);
            string extFile = fileInfo.Extension;
            string nameFile = fileInfo.Name.Substring(0, fileInfo.Name.IndexOf(extFile));
            string input = VBasic.InputBox("Введите новое имя файла", "Переименование", nameFile, 700, 500);
            if (!input.Equals("")) {
                File.Move(tbWorkFolder.Text + @"\" + oldName, tbWorkFolder.Text + @"\" + input + extFile);
            }
            ListRefresh();
        }

        private void ListRefresh() {
            _volumeAsm.reReadListFile(this.listViewWithReordering, MainFormAsm.iniSet.TbWorkFolder);
        }

        // Обработка документов - вставка текста, нумерация, коррекция и т.д.
        private void btnBuild_Click(object sender, EventArgs e) {
            if (listViewWithReordering.SelectedItems.Count == 0) {
                MessageBox.Show(@"Ни одного файла не выбрано!", @"Внимание!");
                return;
            }
            _volumeAsm.setMapBookmarks(mapBookmak);
            foreach (ListViewItem selectedItem in listViewWithReordering.SelectedItems) {
                string file = tbWorkFolder.Text + @"\" + selectedItem.Text;
                if (File.Exists(file)) {
                    _volumeAsm.RebuildDoc(file);
                }
            }
            ListRefresh();
        }

        private void btnImagePath_Click_1(object sender, EventArgs e) {
            folderBrowserDialog.SelectedPath = txtImagePath.Text;
            folderBrowserDialog.ShowDialog();
            txtImagePath.Text = folderBrowserDialog.SelectedPath;
            SaveConfig();
        }

        private void btnEstimatePathAndName_Click_1(object sender, EventArgs e) {
            folderBrowserDialog.SelectedPath = txtEsimatePath.Text;
            folderBrowserDialog.ShowDialog();
            txtEsimatePath.Text = folderBrowserDialog.SelectedPath;
            SaveConfig();
        }
    }
}
