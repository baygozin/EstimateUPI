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
using Microsoft.Office.Core;

namespace EstimatesAssembly {
    // Обработка всех входящих в книгу документов 
    public class VolumeAsm {

        Dictionary<string, string> mapBook = new Dictionary<string, string>();
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

        public void setMapBookmarks(Dictionary<string, string> dict) {
            mapBook = dict;
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
                        foreach (Word.Bookmark docBookmark in _docWord.Bookmarks) {
                            // Обработаем все закладки
                            string bmark = docBookmark.Name;
                            string[] bmaStrings = bmark.Split('_');
                            if (mapBook.ContainsKey(bmark)) {
                                if (!bmaStrings[0].Contains("подпись")) {
                                    // если метка не является подписью (картинкой)
                                    docBookmark.Range.Text = mapBook[docBookmark.Name];
                                } else { // вставляем картинки
                                    string fio = "";
                                    if (bmaStrings[1].Contains("гип")) {
                                        fio = mapBook["фио_гип"];
                                    } else if (bmaStrings[1].Contains("руководителя")) {
                                        fio = mapBook["фио_руководителя"];
                                    }
                                    InsertImageSign(docBookmark, fio);
                                }
                            }
                        }
                    }
                }
            }
            _docWord.Save();
            _docWord.Close();
            _appWord.Quit();
        }

        // Вставить картинку
        private void InsertImageSign(Word.Bookmark bookmark, string fio) {
            try {
                string fileImage = "";
                string fName1 = MainFormAsm.iniSet.TxtImagePath + @"\" + ConvertNameIOF(fio).ToUpper() + ".jpg";
                string fName2 = MainFormAsm.iniSet.TxtImagePath + @"\" + ConvertNameIOF(fio).ToUpper() + ".tif";
                string fName3 = MainFormAsm.iniSet.TxtImagePath + @"\" + ConvertNameIOF(fio).ToUpper() + ".tiff";
                if (File.Exists(fName1)) {
                    fileImage = fName1;
                } else if (File.Exists(fName2)) {
                    fileImage = fName2;
                } else if (File.Exists(fName3)) {
                    fileImage = fName2;
                }
                if (!fileImage.Equals("")) {
                    var picture = bookmark.Range.InlineShapes.AddPicture(fileImage, false, true);
                    if (picture != null) {
                        picture.Height = 50;
                        picture.Width = 100;

                        picture.PictureFormat.TransparentBackground = MsoTriState.msoTrue;
                        picture.PictureFormat.TransparencyColor = ColorTranslator.ToOle(Color.White);
                        picture.Fill.Visible = MsoTriState.msoFalse;
                        var shape = picture.ConvertToShape();
                        shape.WrapFormat.Type = Word.WdWrapType.wdWrapFront;
                    }
                }
            } catch (Exception e) {
                MessageBox.Show(e.Message, @"Ошибка при работе с изображением!");
            }
        }

        // Преобразовать Ф И.О. к виду Ф_И_О
        private static string ConvertNameIOF(string name) {
            string n = name.Replace(".", "_").Replace(" ", "_");
            n = n.Substring(0, n.Length);
            return n;
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