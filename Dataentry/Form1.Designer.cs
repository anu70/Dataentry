namespace Dataentry
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
            this.BrowseButton = new System.Windows.Forms.Button();
            this.ConvertToExcelButton = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.TextFilePathtextBox = new System.Windows.Forms.TextBox();
            this.TextFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.SaveExcelFileDialog = new System.Windows.Forms.SaveFileDialog();
            this.ProgressBar = new System.Windows.Forms.StatusStrip();
            this.toolStripProgressBar1 = new System.Windows.Forms.ToolStripProgressBar();
            this.ProgressBar.SuspendLayout();
            this.SuspendLayout();
            // 
            // BrowseButton
            // 
            this.BrowseButton.Location = new System.Drawing.Point(268, 69);
            this.BrowseButton.Name = "BrowseButton";
            this.BrowseButton.Size = new System.Drawing.Size(75, 23);
            this.BrowseButton.TabIndex = 0;
            this.BrowseButton.Text = "Browse";
            this.BrowseButton.UseVisualStyleBackColor = true;
            this.BrowseButton.Click += new System.EventHandler(this.BrowseButton_Click);
            // 
            // ConvertToExcelButton
            // 
            this.ConvertToExcelButton.Location = new System.Drawing.Point(88, 108);
            this.ConvertToExcelButton.Name = "ConvertToExcelButton";
            this.ConvertToExcelButton.Size = new System.Drawing.Size(121, 23);
            this.ConvertToExcelButton.TabIndex = 1;
            this.ConvertToExcelButton.Text = "Convert to Excel";
            this.ConvertToExcelButton.UseVisualStyleBackColor = true;
            this.ConvertToExcelButton.Click += new System.EventHandler(this.ConvertToExcelButton_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(84, 35);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(148, 19);
            this.label1.TabIndex = 2;
            this.label1.Text = "Select File To Convert";
            // 
            // TextFilePathtextBox
            // 
            this.TextFilePathtextBox.Location = new System.Drawing.Point(55, 69);
            this.TextFilePathtextBox.Name = "TextFilePathtextBox";
            this.TextFilePathtextBox.Size = new System.Drawing.Size(194, 20);
            this.TextFilePathtextBox.TabIndex = 3;
            // 
            // TextFileDialog
            // 
            this.TextFileDialog.FileName = "TextFileDialog";
            this.TextFileDialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            this.TextFileDialog.Title = "Browse Text Files";
            // 
            // SaveExcelFileDialog
            // 
            this.SaveExcelFileDialog.FileName = "SaveExcelFileDialog";
            this.SaveExcelFileDialog.Title = "Save As";
            // 
            // ProgressBar
            // 
            this.ProgressBar.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripProgressBar1});
            this.ProgressBar.Location = new System.Drawing.Point(0, 520);
            this.ProgressBar.Name = "ProgressBar";
            this.ProgressBar.Size = new System.Drawing.Size(752, 22);
            this.ProgressBar.TabIndex = 4;
            this.ProgressBar.Text = "statusStrip1";
            // 
            // toolStripProgressBar1
            // 
            this.toolStripProgressBar1.Name = "toolStripProgressBar1";
            this.toolStripProgressBar1.Size = new System.Drawing.Size(100, 16);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(752, 542);
            this.Controls.Add(this.ProgressBar);
            this.Controls.Add(this.TextFilePathtextBox);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.ConvertToExcelButton);
            this.Controls.Add(this.BrowseButton);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ProgressBar.ResumeLayout(false);
            this.ProgressBar.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button BrowseButton;
        private System.Windows.Forms.Button ConvertToExcelButton;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox TextFilePathtextBox;
        private System.Windows.Forms.OpenFileDialog TextFileDialog;
        private System.Windows.Forms.SaveFileDialog SaveExcelFileDialog;
        private System.Windows.Forms.StatusStrip ProgressBar;
        private System.Windows.Forms.ToolStripProgressBar toolStripProgressBar1;
    }
}

