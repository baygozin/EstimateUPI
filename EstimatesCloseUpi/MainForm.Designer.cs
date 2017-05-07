namespace EstimatesName {
    partial class MainForm {
        /// <summary>
        /// Требуется переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Обязательный метод для поддержки конструктора - не изменяйте
        /// содержимое данного метода при помощи редактора кода.
        /// </summary>
        private void InitializeComponent() {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.показатьToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.закрытьToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.закрытьССохранениемToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.закрытьВсеToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.закрытьВСЕССохранениемToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.обновитьСписокToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.button2 = new System.Windows.Forms.Button();
            this.tbWorkPath = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnGetPath = new System.Windows.Forms.Button();
            this.dlgPath = new System.Windows.Forms.FolderBrowserDialog();
            this.columnID = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnName = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.listViewProcess = new System.Windows.Forms.ListView();
            this.lblNumTable = new System.Windows.Forms.Label();
            this.contextMenuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.показатьToolStripMenuItem,
            this.закрытьToolStripMenuItem,
            this.закрытьССохранениемToolStripMenuItem,
            this.закрытьВсеToolStripMenuItem,
            this.закрытьВСЕССохранениемToolStripMenuItem,
            this.обновитьСписокToolStripMenuItem});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(230, 136);
            // 
            // показатьToolStripMenuItem
            // 
            this.показатьToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("показатьToolStripMenuItem.Image")));
            this.показатьToolStripMenuItem.Name = "показатьToolStripMenuItem";
            this.показатьToolStripMenuItem.Size = new System.Drawing.Size(229, 22);
            this.показатьToolStripMenuItem.Text = "Показать";
            this.показатьToolStripMenuItem.Click += new System.EventHandler(this.показатьToolStripMenuItem_Click);
            // 
            // закрытьToolStripMenuItem
            // 
            this.закрытьToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("закрытьToolStripMenuItem.Image")));
            this.закрытьToolStripMenuItem.Name = "закрытьToolStripMenuItem";
            this.закрытьToolStripMenuItem.Size = new System.Drawing.Size(229, 22);
            this.закрытьToolStripMenuItem.Text = "Закрыть";
            this.закрытьToolStripMenuItem.Click += new System.EventHandler(this.закрытьToolStripMenuItem_Click);
            // 
            // закрытьССохранениемToolStripMenuItem
            // 
            this.закрытьССохранениемToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("закрытьССохранениемToolStripMenuItem.Image")));
            this.закрытьССохранениемToolStripMenuItem.Name = "закрытьССохранениемToolStripMenuItem";
            this.закрытьССохранениемToolStripMenuItem.Size = new System.Drawing.Size(229, 22);
            this.закрытьССохранениемToolStripMenuItem.Text = "Закрыть с сохранением";
            this.закрытьССохранениемToolStripMenuItem.Click += new System.EventHandler(this.закрытьССохранениемToolStripMenuItem_Click);
            // 
            // закрытьВсеToolStripMenuItem
            // 
            this.закрытьВсеToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("закрытьВсеToolStripMenuItem.Image")));
            this.закрытьВсеToolStripMenuItem.Name = "закрытьВсеToolStripMenuItem";
            this.закрытьВсеToolStripMenuItem.Size = new System.Drawing.Size(229, 22);
            this.закрытьВсеToolStripMenuItem.Text = "Закрыть ВСЕ";
            this.закрытьВсеToolStripMenuItem.Click += new System.EventHandler(this.закрытьВсеToolStripMenuItem_Click);
            // 
            // закрытьВСЕССохранениемToolStripMenuItem
            // 
            this.закрытьВСЕССохранениемToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("закрытьВСЕССохранениемToolStripMenuItem.Image")));
            this.закрытьВСЕССохранениемToolStripMenuItem.Name = "закрытьВСЕССохранениемToolStripMenuItem";
            this.закрытьВСЕССохранениемToolStripMenuItem.Size = new System.Drawing.Size(229, 22);
            this.закрытьВСЕССохранениемToolStripMenuItem.Text = "Закрыть ВСЕ с сохранением";
            this.закрытьВСЕССохранениемToolStripMenuItem.Click += new System.EventHandler(this.закрытьВСЕССохранениемToolStripMenuItem_Click);
            // 
            // обновитьСписокToolStripMenuItem
            // 
            this.обновитьСписокToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("обновитьСписокToolStripMenuItem.Image")));
            this.обновитьСписокToolStripMenuItem.Name = "обновитьСписокToolStripMenuItem";
            this.обновитьСписокToolStripMenuItem.Size = new System.Drawing.Size(229, 22);
            this.обновитьСписокToolStripMenuItem.Text = "Обновить список";
            this.обновитьСписокToolStripMenuItem.Click += new System.EventHandler(this.обновитьСписокToolStripMenuItem_Click);
            // 
            // button2
            // 
            this.button2.Image = ((System.Drawing.Image)(resources.GetObject("button2.Image")));
            this.button2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button2.Location = new System.Drawing.Point(805, 543);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(107, 23);
            this.button2.TabIndex = 2;
            this.button2.Text = "Выход";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // tbWorkPath
            // 
            this.tbWorkPath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tbWorkPath.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.tbWorkPath.Location = new System.Drawing.Point(106, 12);
            this.tbWorkPath.Name = "tbWorkPath";
            this.tbWorkPath.Size = new System.Drawing.Size(770, 20);
            this.tbWorkPath.TabIndex = 3;
            this.tbWorkPath.TextChanged += new System.EventHandler(this.tbWorkPath_TextChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 14);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(88, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "Рабочая папка :";
            // 
            // btnGetPath
            // 
            this.btnGetPath.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnGetPath.Image = ((System.Drawing.Image)(resources.GetObject("btnGetPath.Image")));
            this.btnGetPath.Location = new System.Drawing.Point(882, 12);
            this.btnGetPath.Name = "btnGetPath";
            this.btnGetPath.Size = new System.Drawing.Size(29, 22);
            this.btnGetPath.TabIndex = 5;
            this.btnGetPath.UseVisualStyleBackColor = true;
            this.btnGetPath.Click += new System.EventHandler(this.button3_Click);
            // 
            // dlgPath
            // 
            this.dlgPath.Description = "Рабочая папка";
            this.dlgPath.RootFolder = System.Environment.SpecialFolder.MyComputer;
            // 
            // columnID
            // 
            this.columnID.Text = "ID";
            this.columnID.Width = 80;
            // 
            // columnName
            // 
            this.columnName.Text = "Открытые документы Excel";
            this.columnName.Width = 819;
            // 
            // listViewProcess
            // 
            this.listViewProcess.Alignment = System.Windows.Forms.ListViewAlignment.SnapToGrid;
            this.listViewProcess.AutoArrange = false;
            this.listViewProcess.BackColor = System.Drawing.SystemColors.Window;
            this.listViewProcess.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnID,
            this.columnName});
            this.listViewProcess.ContextMenuStrip = this.contextMenuStrip1;
            this.listViewProcess.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.listViewProcess.FullRowSelect = true;
            this.listViewProcess.GridLines = true;
            this.listViewProcess.HideSelection = false;
            this.listViewProcess.Location = new System.Drawing.Point(12, 38);
            this.listViewProcess.MultiSelect = false;
            this.listViewProcess.Name = "listViewProcess";
            this.listViewProcess.ShowGroups = false;
            this.listViewProcess.Size = new System.Drawing.Size(900, 500);
            this.listViewProcess.Sorting = System.Windows.Forms.SortOrder.Ascending;
            this.listViewProcess.TabIndex = 1;
            this.listViewProcess.UseCompatibleStateImageBehavior = false;
            this.listViewProcess.View = System.Windows.Forms.View.Details;
            this.listViewProcess.ColumnClick += new System.Windows.Forms.ColumnClickEventHandler(this.listView1_ColumnClick);
            // 
            // lblNumTable
            // 
            this.lblNumTable.AutoSize = true;
            this.lblNumTable.Location = new System.Drawing.Point(9, 548);
            this.lblNumTable.Name = "lblNumTable";
            this.lblNumTable.Size = new System.Drawing.Size(35, 13);
            this.lblNumTable.TabIndex = 29;
            this.lblNumTable.Text = "label9";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlLight;
            this.ClientSize = new System.Drawing.Size(923, 572);
            this.Controls.Add(this.lblNumTable);
            this.Controls.Add(this.btnGetPath);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.tbWorkPath);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.listViewProcess);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Сметы";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.MainForm_FormClosing);
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.Shown += new System.EventHandler(this.MainForm_Shown);
            this.contextMenuStrip1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem показатьToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem закрытьToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem закрытьВсеToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem закрытьССохранениемToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem закрытьВСЕССохранениемToolStripMenuItem;
        private System.Windows.Forms.TextBox tbWorkPath;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnGetPath;
        private System.Windows.Forms.FolderBrowserDialog dlgPath;
        private System.Windows.Forms.ToolStripMenuItem обновитьСписокToolStripMenuItem;
        private System.Windows.Forms.ColumnHeader columnID;
        private System.Windows.Forms.ColumnHeader columnName;
        private System.Windows.Forms.ListView listViewProcess;
        private System.Windows.Forms.Label lblNumTable;
    }
}

