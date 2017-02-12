namespace ExcelProjectV2
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.قاعدةالبياناتToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.تحديثقواعدالبياناتToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.lbl_Path = new System.Windows.Forms.Label();
            this.btn_importExcel = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.button1 = new System.Windows.Forms.Button();
            this.dg_data = new System.Windows.Forms.DataGridView();
            this.lb_sheets = new System.Windows.Forms.ListBox();
            this.btn_Save = new System.Windows.Forms.Button();
            this.OFD_Excel = new System.Windows.Forms.OpenFileDialog();
            this.label1 = new System.Windows.Forms.Label();
            this.txt_RowNum = new System.Windows.Forms.NumericUpDown();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.rb_Atm = new System.Windows.Forms.RadioButton();
            this.rb_Bank = new System.Windows.Forms.RadioButton();
            this.menuStrip1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dg_data)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txt_RowNum)).BeginInit();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.Color.Green;
            this.menuStrip1.Font = new System.Drawing.Font("Andalus", 14F);
            this.menuStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.قاعدةالبياناتToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.menuStrip1.Size = new System.Drawing.Size(1191, 46);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // قاعدةالبياناتToolStripMenuItem
            // 
            this.قاعدةالبياناتToolStripMenuItem.BackColor = System.Drawing.Color.DarkGreen;
            this.قاعدةالبياناتToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.تحديثقواعدالبياناتToolStripMenuItem});
            this.قاعدةالبياناتToolStripMenuItem.ForeColor = System.Drawing.Color.Black;
            this.قاعدةالبياناتToolStripMenuItem.Name = "قاعدةالبياناتToolStripMenuItem";
            this.قاعدةالبياناتToolStripMenuItem.Size = new System.Drawing.Size(138, 42);
            this.قاعدةالبياناتToolStripMenuItem.Text = "قاعدة البيانات";
            // 
            // تحديثقواعدالبياناتToolStripMenuItem
            // 
            this.تحديثقواعدالبياناتToolStripMenuItem.Name = "تحديثقواعدالبياناتToolStripMenuItem";
            this.تحديثقواعدالبياناتToolStripMenuItem.Size = new System.Drawing.Size(267, 42);
            this.تحديثقواعدالبياناتToolStripMenuItem.Text = "تحديث قواعد البيانات";
            this.تحديثقواعدالبياناتToolStripMenuItem.Click += new System.EventHandler(this.تحديثقواعدالبياناتToolStripMenuItem_Click);
            // 
            // lbl_Path
            // 
            this.lbl_Path.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_Path.Location = new System.Drawing.Point(0, 77);
            this.lbl_Path.Name = "lbl_Path";
            this.lbl_Path.Size = new System.Drawing.Size(1048, 55);
            this.lbl_Path.TabIndex = 2;
            this.lbl_Path.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // btn_importExcel
            // 
            this.btn_importExcel.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_importExcel.Image = global::ExcelProjectV2.Properties.Resources.Excel_icon_large;
            this.btn_importExcel.Location = new System.Drawing.Point(1054, 60);
            this.btn_importExcel.Name = "btn_importExcel";
            this.btn_importExcel.Size = new System.Drawing.Size(125, 94);
            this.btn_importExcel.TabIndex = 1;
            this.btn_importExcel.TextAlign = System.Drawing.ContentAlignment.TopRight;
            this.btn_importExcel.UseVisualStyleBackColor = true;
            this.btn_importExcel.Click += new System.EventHandler(this.btn_importExcel_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.button1);
            this.groupBox1.Controls.Add(this.dg_data);
            this.groupBox1.Controls.Add(this.lb_sheets);
            this.groupBox1.Location = new System.Drawing.Point(12, 212);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(1175, 453);
            this.groupBox1.TabIndex = 3;
            this.groupBox1.TabStop = false;
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.ForestGreen;
            this.button1.Location = new System.Drawing.Point(915, 187);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(55, 78);
            this.button1.TabIndex = 2;
            this.button1.Text = "<";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // dg_data
            // 
            this.dg_data.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dg_data.Location = new System.Drawing.Point(7, 12);
            this.dg_data.Name = "dg_data";
            this.dg_data.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.dg_data.RowTemplate.Height = 24;
            this.dg_data.Size = new System.Drawing.Size(908, 436);
            this.dg_data.TabIndex = 1;
            // 
            // lb_sheets
            // 
            this.lb_sheets.FormattingEnabled = true;
            this.lb_sheets.ItemHeight = 16;
            this.lb_sheets.Location = new System.Drawing.Point(971, 12);
            this.lb_sheets.Name = "lb_sheets";
            this.lb_sheets.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple;
            this.lb_sheets.Size = new System.Drawing.Size(197, 436);
            this.lb_sheets.TabIndex = 0;
            // 
            // btn_Save
            // 
            this.btn_Save.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_Save.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_Save.Location = new System.Drawing.Point(12, 671);
            this.btn_Save.Name = "btn_Save";
            this.btn_Save.Size = new System.Drawing.Size(1175, 54);
            this.btn_Save.TabIndex = 4;
            this.btn_Save.Text = "حفظ";
            this.btn_Save.UseVisualStyleBackColor = true;
            this.btn_Save.Click += new System.EventHandler(this.btn_Save_Click);
            // 
            // OFD_Excel
            // 
            this.OFD_Excel.DefaultExt = "\"xls\"";
            this.OFD_Excel.Filter = "\"txt files (*.txt)|*.txt|All files (*.*)|*.*\"";
            this.OFD_Excel.RestoreDirectory = true;
            this.OFD_Excel.Title = "Browse Excel Files";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(1060, 184);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(119, 25);
            this.label1.TabIndex = 5;
            this.label1.Text = "سطر الأكواد رقم";
            // 
            // txt_RowNum
            // 
            this.txt_RowNum.Location = new System.Drawing.Point(807, 188);
            this.txt_RowNum.Name = "txt_RowNum";
            this.txt_RowNum.Size = new System.Drawing.Size(120, 22);
            this.txt_RowNum.TabIndex = 6;
            this.txt_RowNum.Value = new decimal(new int[] {
            6,
            0,
            0,
            0});
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.rb_Atm);
            this.groupBox2.Controls.Add(this.rb_Bank);
            this.groupBox2.Location = new System.Drawing.Point(243, 176);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.groupBox2.Size = new System.Drawing.Size(336, 34);
            this.groupBox2.TabIndex = 7;
            this.groupBox2.TabStop = false;
            // 
            // rb_Atm
            // 
            this.rb_Atm.AutoSize = true;
            this.rb_Atm.Checked = true;
            this.rb_Atm.Location = new System.Drawing.Point(196, 9);
            this.rb_Atm.Name = "rb_Atm";
            this.rb_Atm.Size = new System.Drawing.Size(97, 21);
            this.rb_Atm.TabIndex = 0;
            this.rb_Atm.TabStop = true;
            this.rb_Atm.Text = "بطاقات حكومية";
            this.rb_Atm.UseVisualStyleBackColor = true;
            // 
            // rb_Bank
            // 
            this.rb_Bank.AutoSize = true;
            this.rb_Bank.Location = new System.Drawing.Point(18, 9);
            this.rb_Bank.Name = "rb_Bank";
            this.rb_Bank.Size = new System.Drawing.Size(91, 21);
            this.rb_Bank.TabIndex = 0;
            this.rb_Bank.TabStop = true;
            this.rb_Bank.Text = "تحويلات بنكية";
            this.rb_Bank.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(1191, 737);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.txt_RowNum);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btn_Save);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.lbl_Path);
            this.Controls.Add(this.btn_importExcel);
            this.Controls.Add(this.menuStrip1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MainMenuStrip = this.menuStrip1;
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.Text = "برنامج أكواد الدفع الألكترونى - كلية الطب";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dg_data)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txt_RowNum)).EndInit();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem قاعدةالبياناتToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem تحديثقواعدالبياناتToolStripMenuItem;
        private System.Windows.Forms.Button btn_importExcel;
        private System.Windows.Forms.Label lbl_Path;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.DataGridView dg_data;
        private System.Windows.Forms.ListBox lb_sheets;
        private System.Windows.Forms.Button btn_Save;
        private System.Windows.Forms.OpenFileDialog OFD_Excel;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.NumericUpDown txt_RowNum;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.RadioButton rb_Atm;
        private System.Windows.Forms.RadioButton rb_Bank;
    }
}

