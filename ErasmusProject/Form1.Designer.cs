namespace ErasmusProject
{
    partial class ErasmusProject
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
            this.button1 = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.lblTitle = new System.Windows.Forms.Label();
            this.btnLoadExcel = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.lblTimer = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("Cooper Std Black", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.Location = new System.Drawing.Point(17, 229);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(200, 49);
            this.button1.TabIndex = 0;
            this.button1.Text = "START";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(11, 64);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(443, 23);
            this.progressBar1.TabIndex = 1;
            // 
            // listBox1
            // 
            this.listBox1.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.listBox1.FormattingEnabled = true;
            this.listBox1.ItemHeight = 18;
            this.listBox1.Location = new System.Drawing.Point(12, 93);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(442, 130);
            this.listBox1.TabIndex = 2;
            // 
            // lblTitle
            // 
            this.lblTitle.AutoSize = true;
            this.lblTitle.Font = new System.Drawing.Font("Miramonte", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.Location = new System.Drawing.Point(12, 22);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(239, 26);
            this.lblTitle.TabIndex = 3;
            this.lblTitle.Text = "MySql to Excel importer";
            // 
            // btnLoadExcel
            // 
            this.btnLoadExcel.Font = new System.Drawing.Font("Cooper Std Black", 9.749999F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnLoadExcel.Location = new System.Drawing.Point(254, 229);
            this.btnLoadExcel.Name = "btnLoadExcel";
            this.btnLoadExcel.Size = new System.Drawing.Size(200, 49);
            this.btnLoadExcel.TabIndex = 4;
            this.btnLoadExcel.Text = "Load excel document";
            this.btnLoadExcel.UseVisualStyleBackColor = true;
            this.btnLoadExcel.Click += new System.EventHandler(this.btnLoadExcel_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // lblTimer
            // 
            this.lblTimer.Font = new System.Drawing.Font("Adobe Arabic", 24F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTimer.Location = new System.Drawing.Point(254, 9);
            this.lblTimer.Name = "lblTimer";
            this.lblTimer.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.lblTimer.Size = new System.Drawing.Size(200, 52);
            this.lblTimer.TabIndex = 5;
            this.lblTimer.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // ErasmusProject
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(466, 372);
            this.Controls.Add(this.lblTimer);
            this.Controls.Add(this.btnLoadExcel);
            this.Controls.Add(this.lblTitle);
            this.Controls.Add(this.listBox1);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.button1);
            this.Name = "ErasmusProject";
            this.Text = "ErasmusProject";
            this.TopMost = true;
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.ListBox listBox1;
        private System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.Button btnLoadExcel;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Label lblTimer;
    }
}

