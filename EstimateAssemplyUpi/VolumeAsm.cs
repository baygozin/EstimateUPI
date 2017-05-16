using System;
using System.Collections;
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

        Dictionary<string, string> fields = new Dictionary<string, string>();

        public Word._Application _appWord;
        public Word._Document _docWord;
        public Excel._Application _appExcel;
        public Excel._Workbook _excelWorkbook;
        public Excel._Worksheet _excelWorksheet;
        Object _missingObj = System.Reflection.Missing.Value;
        Object trueObj = true;
        Object falseObj = false;
        private string propertyDocName = "Comments";
        private string propertyExcelDocValue = "собранная книга со сметами";
        private string propertyWordDocTitleAll = "титул генпроектировщика";
        private string propertyWordDocTitleUpi = "титул юпи";
        private string propertyWordDocOgl = "оглавление";
        private string propertyWordDocPz = "пояснительная записка";

        public VolumeAsm() {
        }

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
            _docWord = _appWord.Documents.Open(file);
            object val = GetWordDocumentProperty(propertyDocName);
            if (val != null) {
                string propValue = val as string;
                string[] prop = propValue.Split(' ');
                if (prop[0].Equals("шаблон")) {
                    if (prop[1].Equals("титул")) {
                        string spisok = "";
                        foreach (Word.Bookmark documentBookmark in _docWord.Bookmarks) {
                            //documentBookmark.Range.Text = "";
                            // Обработаем все закладки
                            spisok = spisok + documentBookmark.Name + "\n\r";
                        }
                        MessageBox.Show(spisok, "=============");
                    }
                }
            }
            _docWord.Close();
            _appWord.Quit();
        }

        public void RebuildDocExcel(string file) {
            //создаем обьект приложения excel
            _appExcel = new Excel.Application() { Visible = false, DisplayAlerts = false } ;
            _excelWorkbook = _appExcel.Workbooks.Open(file);
            if (_appExcel.Workbooks.Count == 0) {
                MessageBox.Show("Эта книга пустая!", "Внимание!");
                _appExcel.Quit();
                return;
            }
            _excelWorkbook = _appExcel.ActiveWorkbook;
            object val = GetExcelDocumentProperty(propertyDocName);
            if (val != null) {
                string propValue = val as string;
                if (propValue.Equals(propertyExcelDocValue)) {
                    //MessageBox.Show(@"Это наша книга!", "Урааааа!!!");
                } 
            } 
            _excelWorkbook.Close();
            _appExcel.Quit();
        }

        public object GetWordDocumentProperty(string propertyName) {
            object returnVal = null;
            if (_appWord.ActiveDocument == null) {
                return returnVal;
            }
            dynamic properties = _appWord.ActiveDocument.BuiltInDocumentProperties;
            foreach (dynamic property in properties) {
                string name = property.Name;
                if (name.Equals(propertyName)) {
                    returnVal = property.Value;
                }
            }
            return returnVal;
        }

        private object GetExcelDocumentProperty(string propertyName) {
            object returnVal = null;
            if (_appExcel.ActiveWorkbook == null) {
                return returnVal;
            }
            dynamic properties = _appExcel.ActiveWorkbook.BuiltinDocumentProperties;
            foreach (dynamic property in properties) {
                string name = property.Name;
                if (name.Equals(propertyName)) {
                    returnVal = property.Value;
                }
            }
            return returnVal;
        }

    }
}