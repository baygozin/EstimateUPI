using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace AsmBook
{
    public partial class FormAsm : Form
    {
        private string[] fileList;
        public string loadPath;
        public static ConfigApp ini = new ConfigApp();
        private string configfile;

        public FormAsm()
        {
            InitializeComponent();
            configfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\bookasm.xml";
            ReadConfig();
        }

        private void FormAsm_FormClosed(object sender, FormClosedEventArgs e)
        {
            SaveConfig();
        }

        private void ReadConfig()
        {
            if (File.Exists(configfile))
            {
                using (Stream stream = new FileStream(configfile, FileMode.Open))
                {
                    var serializer = new XmlSerializer(typeof(ConfigApp));
                    ini = (ConfigApp) serializer.Deserialize(stream);

                }
                loadPath = ini.loadPath;
            }
        }

        private void SaveConfig()
        {
            ini.loadPath = loadPath;
            using (Stream writer = new FileStream(configfile, FileMode.Create))
            {
                var serializer = new XmlSerializer(typeof(ConfigApp));
                serializer.Serialize(writer, ini);
            }
        }

        private void buttonAdd_Click(object sender, EventArgs e)
        {
            openFileDialogDoc.InitialDirectory = loadPath;
            openFileDialogDoc.CheckFileExists = true;
            openFileDialogDoc.ShowDialog();
            fileList = openFileDialogDoc.FileNames;
            foreach (string file in fileList)
            {
                ListViewItem item = new ListViewItem(new string[] {"1234", file});
                listViewDoc.Items.Add(item);
            }
            loadPath = openFileDialogDoc.InitialDirectory;
        }

        private void listViewDoc_ItemDrag(object sender, ItemDragEventArgs e)
        {
            //Begins a drag-and-drop operation in the ListView 
            listViewDoc.DoDragDrop(listViewDoc.SelectedItems, DragDropEffects.Move);
        }

        private void listViewDoc_DragEnter(object sender, DragEventArgs e)
        {
            int len = e.Data.GetFormats().Length - 1; int i;
            for (i = 0; i <= len; i++)
            {
                if (e.Data.GetFormats()[i].Equals("System.Windows.Forms.ListView+SelectedListViewItemCollection"))
                {
                    //The data from the drag source is moved to the target.
                    e.Effect = DragDropEffects.Move;
                }
            }
        }

        private void listViewDoc_DragDrop(object sender, DragEventArgs e)
        {
            //Return if the items are not selected in the ListView 
            if (listViewDoc.SelectedItems.Count==0)
            {
                return;
            }
            //Returns the location of the mouse pointer in the ListView control.
            Point cp = listViewDoc.PointToClient(new Point(e.X, e.Y));
            //Obtain the item that is located at the specified location of the mouse pointer.
            ListViewItem dragToItem = listViewDoc.GetItemAt(cp.X, cp.Y);
            if (dragToItem == null)
            {
                return;
            }
            //Obtain the index of the item at the mouse pointer.
            int dragIndex = dragToItem.Index;
            ListViewItem[] sel = new ListViewItem[listViewDoc.SelectedItems.Count];
            for (int i = 0; i <= listViewDoc.SelectedItems.Count - 1; i++)
            {
                sel[i] = listViewDoc.SelectedItems[i];
            }
            for (int i = 0; i < sel.GetLength(0); i++)
            {
                //Obtain the ListViewItem to be dragged to the target location.
                ListViewItem dragItem = sel[i];
                int itemIndex = dragIndex;
                if (itemIndex == dragItem.Index)
                {
                    return;
                }
                if (dragItem.Index < itemIndex)
                    itemIndex++;
                else
                    itemIndex = dragIndex + i;
                //Insert the item at the mouse pointer.
                ListViewItem insertItem = (ListViewItem)dragItem.Clone();
                listViewDoc.Items.Insert(itemIndex, insertItem);
                //Removes the item from the initial location while 
                //the item is moved to the new location.
                listViewDoc.Items.Remove(dragItem);
            }

        }

        private void buttonDelete_Click(object sender, EventArgs e)
        {
            if (listViewDoc.SelectedItems.Count == 0)return;
            //ListViewItem[] sel = new ListViewItem[listViewDoc.SelectedItems.Count];
            for (int i = 0; i <= listViewDoc.SelectedItems.Count - 1; i++)
            {
                listViewDoc.SelectedItems[i].Remove();
            }
        }

    }
}
