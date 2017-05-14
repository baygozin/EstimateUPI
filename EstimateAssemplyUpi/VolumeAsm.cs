using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace EstimatesAssembly {
    public class VolumeAsm {
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
                        case ".dotx": {
                            item.BackColor = Color.BlueViolet;
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
    }
}