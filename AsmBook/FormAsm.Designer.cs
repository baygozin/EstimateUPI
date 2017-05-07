namespace AsmBook
{
    partial class FormAsm
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.ColumnHeader columnName;
            System.Windows.Forms.ColumnHeader columnFiles;
            this.buttonAdd = new System.Windows.Forms.Button();
            this.buttonDelete = new System.Windows.Forms.Button();
            this.openFileDialogDoc = new System.Windows.Forms.OpenFileDialog();
            this.listViewDoc = new System.Windows.Forms.ListView();
            columnName = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            columnFiles = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.SuspendLayout();
            // 
            // columnName
            // 
            columnName.Text = "Документ";
            columnName.Width = 200;
            // 
            // columnFiles
            // 
            columnFiles.Text = "Файл";
            columnFiles.Width = 270;
            // 
            // buttonAdd
            // 
            this.buttonAdd.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonAdd.Location = new System.Drawing.Point(493, 12);
            this.buttonAdd.Name = "buttonAdd";
            this.buttonAdd.Size = new System.Drawing.Size(75, 23);
            this.buttonAdd.TabIndex = 1;
            this.buttonAdd.Text = "Добавить";
            this.buttonAdd.UseVisualStyleBackColor = true;
            this.buttonAdd.Click += new System.EventHandler(this.buttonAdd_Click);
            // 
            // buttonDelete
            // 
            this.buttonDelete.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonDelete.Location = new System.Drawing.Point(493, 41);
            this.buttonDelete.Name = "buttonDelete";
            this.buttonDelete.Size = new System.Drawing.Size(75, 23);
            this.buttonDelete.TabIndex = 2;
            this.buttonDelete.Text = "Удалить";
            this.buttonDelete.UseVisualStyleBackColor = true;
            this.buttonDelete.Click += new System.EventHandler(this.buttonDelete_Click);
            // 
            // openFileDialogDoc
            // 
            this.openFileDialogDoc.InitialDirectory = "C:\\tmp";
            this.openFileDialogDoc.Multiselect = true;
            this.openFileDialogDoc.SupportMultiDottedExtensions = true;
            // 
            // listViewDoc
            // 
            this.listViewDoc.Activation = System.Windows.Forms.ItemActivation.OneClick;
            this.listViewDoc.AllowDrop = true;
            this.listViewDoc.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.listViewDoc.AutoArrange = false;
            this.listViewDoc.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            columnName,
            columnFiles});
            this.listViewDoc.FullRowSelect = true;
            this.listViewDoc.GridLines = true;
            this.listViewDoc.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.listViewDoc.HideSelection = false;
            this.listViewDoc.LabelWrap = false;
            this.listViewDoc.Location = new System.Drawing.Point(12, 12);
            this.listViewDoc.Name = "listViewDoc";
            this.listViewDoc.Size = new System.Drawing.Size(475, 554);
            this.listViewDoc.TabIndex = 3;
            this.listViewDoc.UseCompatibleStateImageBehavior = false;
            this.listViewDoc.View = System.Windows.Forms.View.Details;
            this.listViewDoc.ItemDrag += new System.Windows.Forms.ItemDragEventHandler(this.listViewDoc_ItemDrag);
            this.listViewDoc.DragDrop += new System.Windows.Forms.DragEventHandler(this.listViewDoc_DragDrop);
            this.listViewDoc.DragEnter += new System.Windows.Forms.DragEventHandler(this.listViewDoc_DragEnter);
            // 
            // FormAsm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(576, 578);
            this.Controls.Add(this.listViewDoc);
            this.Controls.Add(this.buttonDelete);
            this.Controls.Add(this.buttonAdd);
            this.Name = "FormAsm";
            this.Text = "Сборка документации";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.FormAsm_FormClosed);
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button buttonAdd;
        private System.Windows.Forms.Button buttonDelete;
        private System.Windows.Forms.OpenFileDialog openFileDialogDoc;
        private System.Windows.Forms.ListView listViewDoc;
    }
}

