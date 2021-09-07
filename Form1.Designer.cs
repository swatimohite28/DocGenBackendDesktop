
namespace export_html_to_word
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
            this.fileDialog = new System.Windows.Forms.OpenFileDialog();
            this.btnExport = new System.Windows.Forms.Button();
            this.btnOpenFile = new System.Windows.Forms.Button();
            this.lblFile = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // fileDialog
            // 
            this.fileDialog.DefaultExt = "html";
            this.fileDialog.Filter = "html files (*.html)|*.html|All files (*.*)|*.*";
            this.fileDialog.FilterIndex = 2;
            this.fileDialog.InitialDirectory = "C:\\";
            this.fileDialog.RestoreDirectory = true;
            this.fileDialog.Title = "Browse html Files";
            // 
            // btnExport
            // 
            this.btnExport.Location = new System.Drawing.Point(312, 156);
            //this.btnExport.Margin = new System.Windows.Forms.Padding(1);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(90, 37);
            this.btnExport.TabIndex = 2;
            this.btnExport.Text = "Save as Word";
            //this.btnExport.UseVisualStyleBackColor = false;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // btnOpenFile
            // 
            this.btnOpenFile.Location = new System.Drawing.Point(312, 28);
            //this.btnOpenFile.Margin = new System.Windows.Forms.Padding(2);
            this.btnOpenFile.Name = "btnOpenFile";
            this.btnOpenFile.Size = new System.Drawing.Size(90, 39);
            this.btnOpenFile.TabIndex = 1;
            this.btnOpenFile.Text = "Open Html File";
            this.btnOpenFile.Click += new System.EventHandler(this.btnOpenFile_Click);
            // 
            // lblFile
            // 
            this.lblFile.Location = new System.Drawing.Point(8, 28);
            this.lblFile.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblFile.Name = "lblFile";
            this.lblFile.Size = new System.Drawing.Size(300, 52);
            this.lblFile.TabIndex = 2;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.ClientSize = new System.Drawing.Size(403, 195);
            this.Controls.Add(this.btnOpenFile);
            this.Controls.Add(this.lblFile);
            this.Controls.Add(this.btnExport);
            this.ForeColor = System.Drawing.SystemColors.ControlText;
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "Form1";
            this.Text = "Document Generation";
            this.TransparencyKey = System.Drawing.Color.Transparent;
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.Button btnOpenFile;
        private System.Windows.Forms.Label lblFile;
        private System.Windows.Forms.OpenFileDialog fileDialog;
    }
}

