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
            this.menuStrip1.SuspendLayout();
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
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(1191, 676);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Form1";
            this.Text = "Form1";
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem قاعدةالبياناتToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem تحديثقواعدالبياناتToolStripMenuItem;
    }
}

