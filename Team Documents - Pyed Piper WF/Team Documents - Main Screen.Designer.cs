namespace Team_Documents___Pyed_Piper_WF
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.panel1 = new System.Windows.Forms.Panel();
            this.lblLoggedOnUser = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.panel2 = new System.Windows.Forms.Panel();
            this.btnProcessFile = new System.Windows.Forms.Button();
            this.btnChooseTheme = new System.Windows.Forms.Button();
            this.btnUploadFile = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.txtFileName = new System.Windows.Forms.TextBox();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.lblProcessing = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.AutoSize = true;
            this.panel1.Controls.Add(this.lblLoggedOnUser);
            this.panel1.Controls.Add(this.button1);
            this.panel1.Location = new System.Drawing.Point(206, 5);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1005, 89);
            this.panel1.TabIndex = 0;
            // 
            // lblLoggedOnUser
            // 
            this.lblLoggedOnUser.AutoSize = true;
            this.lblLoggedOnUser.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblLoggedOnUser.Location = new System.Drawing.Point(788, 62);
            this.lblLoggedOnUser.Name = "lblLoggedOnUser";
            this.lblLoggedOnUser.Size = new System.Drawing.Size(94, 14);
            this.lblLoggedOnUser.TabIndex = 2;
            this.lblLoggedOnUser.Text = "Logged On As : ";
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button1.Location = new System.Drawing.Point(3, 84);
            this.button1.Margin = new System.Windows.Forms.Padding(0);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(1002, 5);
            this.button1.TabIndex = 1;
            this.button1.Text = "button1";
            this.button1.UseVisualStyleBackColor = false;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::Team_Documents___Pyed_Piper_WF.Properties.Resources.New_Stanlib_Logo;
            this.pictureBox1.Location = new System.Drawing.Point(-2, 5);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(241, 99);
            this.pictureBox1.TabIndex = 1;
            this.pictureBox1.TabStop = false;
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.btnProcessFile);
            this.panel2.Controls.Add(this.btnChooseTheme);
            this.panel2.Controls.Add(this.btnUploadFile);
            this.panel2.Controls.Add(this.label1);
            this.panel2.Controls.Add(this.txtFileName);
            this.panel2.Controls.Add(this.comboBox1);
            this.panel2.Location = new System.Drawing.Point(203, 174);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(870, 171);
            this.panel2.TabIndex = 3;
            // 
            // btnProcessFile
            // 
            this.btnProcessFile.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnProcessFile.Location = new System.Drawing.Point(648, 89);
            this.btnProcessFile.Name = "btnProcessFile";
            this.btnProcessFile.Size = new System.Drawing.Size(159, 33);
            this.btnProcessFile.TabIndex = 5;
            this.btnProcessFile.Text = "Process File";
            this.btnProcessFile.UseVisualStyleBackColor = true;
            this.btnProcessFile.Click += new System.EventHandler(this.btnProcessFile_Click);
            // 
            // btnChooseTheme
            // 
            this.btnChooseTheme.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnChooseTheme.Location = new System.Drawing.Point(56, 89);
            this.btnChooseTheme.Name = "btnChooseTheme";
            this.btnChooseTheme.Size = new System.Drawing.Size(149, 33);
            this.btnChooseTheme.TabIndex = 3;
            this.btnChooseTheme.Text = "Choose Theme";
            this.btnChooseTheme.UseVisualStyleBackColor = true;
            // 
            // btnUploadFile
            // 
            this.btnUploadFile.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnUploadFile.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnUploadFile.Location = new System.Drawing.Point(56, 50);
            this.btnUploadFile.Name = "btnUploadFile";
            this.btnUploadFile.Size = new System.Drawing.Size(149, 32);
            this.btnUploadFile.TabIndex = 2;
            this.btnUploadFile.Text = "Select File";
            this.btnUploadFile.UseVisualStyleBackColor = true;
            this.btnUploadFile.Click += new System.EventHandler(this.btnUploadFile_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Arial", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(207, 15);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(545, 22);
            this.label1.TabIndex = 0;
            this.label1.Text = "PLEASE SELECT THE DOCUMENT YOU WISH TO FORMAT:";
            // 
            // txtFileName
            // 
            this.txtFileName.Font = new System.Drawing.Font("Arial", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtFileName.Location = new System.Drawing.Point(211, 50);
            this.txtFileName.Multiline = true;
            this.txtFileName.Name = "txtFileName";
            this.txtFileName.Size = new System.Drawing.Size(596, 32);
            this.txtFileName.TabIndex = 1;
            // 
            // comboBox1
            // 
            this.comboBox1.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Items.AddRange(new object[] {
            "STANLIB BRAND (FULL i.e. with Headers , Footers & TOC)",
            "STANLIB BRAND (PARTIAL)",
            "LIBERTY BRAND  (FULL i.e. with Headers , Footers & TOC)",
            "LIBERTY BRAND (PARTIAL)"});
            this.comboBox1.Location = new System.Drawing.Point(211, 89);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(431, 26);
            this.comboBox1.TabIndex = 4;
            // 
            // lblProcessing
            // 
            this.lblProcessing.AutoSize = true;
            this.lblProcessing.Font = new System.Drawing.Font("Arial", 24F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblProcessing.Location = new System.Drawing.Point(453, 363);
            this.lblProcessing.Name = "lblProcessing";
            this.lblProcessing.Size = new System.Drawing.Size(346, 36);
            this.lblProcessing.TabIndex = 4;
            this.lblProcessing.Text = "Processing document...";
            this.lblProcessing.Visible = false;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1265, 569);
            this.Controls.Add(this.lblProcessing);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.panel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "Team Documents - Pyed Piper Desktop Version";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.Form1_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Button btnUploadFile;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtFileName;
        private System.Windows.Forms.Button btnProcessFile;
        private System.Windows.Forms.Button btnChooseTheme;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.Label lblProcessing;
        private System.Windows.Forms.Label lblLoggedOnUser;
    }
}

