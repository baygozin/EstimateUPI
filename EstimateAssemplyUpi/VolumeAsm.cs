using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using DocumentFormat.OpenXml.Packaging;

namespace EstimatesAssembly {
    public class VolumeAsm {

        public Word._Application _appWord;
        public Word._Document _docWord;
        public Excel._Application _appExcel;
        public Excel._Workbook _excelWorkbook;
        public Excel._Worksheet _excelWorksheet;
        Object _missingObj = System.Reflection.Missing.Value;
        Object trueObj = true;
        Object falseObj = false;
        private string propertyDocName = "Comments";
        private string propertyDocValue = "собранная книга со сметами";



        public void reReadListFile(ListViewWithReordering listView, string folder) {
            if (Directory.Exists(folder)) {
                string[] files = Directory.GetFiles(folder);
                listView.Clear();
                foreach (var file in files) {
                    FileInfo fileInfo = new FileInfo(file);
                    ListViewItem item = new ListViewItem();
                    item.Text = fileInfo.Name;
                    switch (fileInfo.Extension.ToLower()) {
                        case ".docx": {
                            item.BackColor = Color.Aqua;
                            break;
                        }
                        case ".doc": {
                            item.BackColor = Color.Aqua;
                            break;
                        }
                        case ".xlsx": {
                            item.BackColor = Color.Chartreuse;
                            break;
                        }
                        case ".xls": {
                            item.BackColor = Color.Chartreuse;
                            break;
                        }
                        case ".pdf": {
                            item.BackColor = Color.LightCoral;
                            break;
                        }
                        default: {
                            break;
                        }
                    }
                    listView.Items.Add(item);
                }
            }
        }

        public void RebuildDoc(string file) {
            // Определим, что же за тип документа нам втюхивают...
            // А опредклять будем по расширению...
            FileInfo fileInfo = new FileInfo(file);
            if (fileInfo.Extension.Equals(".docx") || fileInfo.Extension.Equals(".doc")) {
                // это документ Word
                RebuildDocWord(file);
            }
            else if (fileInfo.Extension.Equals(".xlsx") || fileInfo.Extension.Equals(".xls")) {
                // это документ Excel
                RebuildDocExcel(file);
            }
            else {
                // А тут все остальное
            }
        }

        public void RebuildDocWord(string file) {
            //создаем обьект приложения word
            _appWord = new Word.Application() { Visible = false };
            // создаем путь к файлу
            Object templatePathObj = file;

            _docWord = _appWord.Documents.Add(ref templatePathObj, ref _missingObj, ref _missingObj, ref _missingObj);
            foreach (Word.Bookmark documentBookmark in _docWord.Bookmarks) {
                //documentBookmark.Range.Text = "";
                // Обработаем все закладки
            }
            _docWord.Close();
            _appWord.Quit();
        }

        public void RebuildDocExcel(string file) {
            _appExcel = new Excel.Application() { Visible = false, DisplayAlerts = false } ;
            _excelWorkbook = _appExcel.Workbooks.Open(file);
            if (_appExcel.Workbooks.Count == 0) {
                MessageBox.Show("Эта книга пустая!", "Внимание!");
                _appExcel.Quit();
                return;
            }
            _excelWorkbook = _appExcel.ActiveWorkbook;
            string propValue = GetDocumentProperty(propertyDocName) as string;
            if (propValue.Equals(propertyDocValue)) {
                //MessageBox.Show(@"Это наша книга!", "Урааааа!!!");
            } else {
                MessageBox.Show(@"Это не наша книга!", "Внимание!");
            }
            _excelWorkbook.Close();
            _appExcel.Quit();
        }

        private object GetDocumentProperty(string propertyName) {
            object returnVal = null;
            dynamic properties = _appExcel.ActiveWorkbook.BuiltinDocumentProperties;
            foreach (dynamic property in properties) {
                string name = property.Name;
                if (name.Equals(propertyName)) {
                    returnVal = property.Value;
                }
            }
            string test = returnVal.ToString();
            return returnVal;
        }

    }
}