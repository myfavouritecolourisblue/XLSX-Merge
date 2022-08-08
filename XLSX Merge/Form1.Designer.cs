namespace XLSX_Merge
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnReadXlsx = new System.Windows.Forms.Button();
            this.lblXlsxFileView = new System.Windows.Forms.Label();
            this.txtbxXlsx = new System.Windows.Forms.TextBox();
            this.lblXlsxFile = new System.Windows.Forms.Label();
            this.txtbxXlsxFile = new System.Windows.Forms.TextBox();
            this.lblWorksheet = new System.Windows.Forms.Label();
            this.lblCell = new System.Windows.Forms.Label();
            this.txtbxCell = new System.Windows.Forms.TextBox();
            this.txtbxCellPicker = new System.Windows.Forms.TextBox();
            this.txtbxWorksheet = new System.Windows.Forms.TextBox();
            this.btnReadCell = new System.Windows.Forms.Button();
            this.btnReadCellRange = new System.Windows.Forms.Button();
            this.txtbxCellRange = new System.Windows.Forms.TextBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.btnManipulateCell = new System.Windows.Forms.Button();
            this.lblCellManipulationNewValue = new System.Windows.Forms.Label();
            this.lblCellManipulationOldValue = new System.Windows.Forms.Label();
            this.txtbxCellManipulationCellPicker = new System.Windows.Forms.TextBox();
            this.txtbxCellManipulationNewValue = new System.Windows.Forms.TextBox();
            this.txtbxCellManipulationOldValue = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.lblManipulateCell = new System.Windows.Forms.Label();
            this.btnSaveXlsx = new System.Windows.Forms.Button();
            this.btnReadCsv = new System.Windows.Forms.Button();
            this.lblCsvFile = new System.Windows.Forms.Label();
            this.txtbxCsvFile = new System.Windows.Forms.TextBox();
            this.txtbxCsvFileView = new System.Windows.Forms.TextBox();
            this.lblCsvFileView = new System.Windows.Forms.Label();
            this.btnMergeFiles = new System.Windows.Forms.Button();
            this.btnConvCsvXlsx = new System.Windows.Forms.Button();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnReadXlsx
            // 
            this.btnReadXlsx.Location = new System.Drawing.Point(12, 12);
            this.btnReadXlsx.Name = "btnReadXlsx";
            this.btnReadXlsx.Size = new System.Drawing.Size(75, 23);
            this.btnReadXlsx.TabIndex = 0;
            this.btnReadXlsx.Text = "Read .xlsx";
            this.btnReadXlsx.UseVisualStyleBackColor = true;
            this.btnReadXlsx.Click += new System.EventHandler(this.btnReadXlsx_Click);
            // 
            // lblXlsxFileView
            // 
            this.lblXlsxFileView.AutoSize = true;
            this.lblXlsxFileView.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.lblXlsxFileView.Location = new System.Drawing.Point(106, 471);
            this.lblXlsxFileView.Name = "lblXlsxFileView";
            this.lblXlsxFileView.Size = new System.Drawing.Size(83, 15);
            this.lblXlsxFileView.TabIndex = 1;
            this.lblXlsxFileView.Text = ".xlsx file view";
            // 
            // txtbxXlsx
            // 
            this.txtbxXlsx.Location = new System.Drawing.Point(12, 489);
            this.txtbxXlsx.Multiline = true;
            this.txtbxXlsx.Name = "txtbxXlsx";
            this.txtbxXlsx.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtbxXlsx.Size = new System.Drawing.Size(291, 210);
            this.txtbxXlsx.TabIndex = 2;
            // 
            // lblXlsxFile
            // 
            this.lblXlsxFile.AutoSize = true;
            this.lblXlsxFile.Location = new System.Drawing.Point(12, 48);
            this.lblXlsxFile.Name = "lblXlsxFile";
            this.lblXlsxFile.Size = new System.Drawing.Size(60, 15);
            this.lblXlsxFile.TabIndex = 3;
            this.lblXlsxFile.Text = ".xlsx Datei";
            // 
            // txtbxXlsxFile
            // 
            this.txtbxXlsxFile.Location = new System.Drawing.Point(78, 45);
            this.txtbxXlsxFile.Name = "txtbxXlsxFile";
            this.txtbxXlsxFile.Size = new System.Drawing.Size(401, 23);
            this.txtbxXlsxFile.TabIndex = 4;
            // 
            // lblWorksheet
            // 
            this.lblWorksheet.AutoSize = true;
            this.lblWorksheet.Location = new System.Drawing.Point(12, 77);
            this.lblWorksheet.Name = "lblWorksheet";
            this.lblWorksheet.Size = new System.Drawing.Size(63, 15);
            this.lblWorksheet.TabIndex = 5;
            this.lblWorksheet.Text = "Worksheet";
            // 
            // lblCell
            // 
            this.lblCell.AutoSize = true;
            this.lblCell.Location = new System.Drawing.Point(40, 106);
            this.lblCell.Name = "lblCell";
            this.lblCell.Size = new System.Drawing.Size(32, 15);
            this.lblCell.TabIndex = 6;
            this.lblCell.Text = "Zelle";
            // 
            // txtbxCell
            // 
            this.txtbxCell.Location = new System.Drawing.Point(78, 103);
            this.txtbxCell.Name = "txtbxCell";
            this.txtbxCell.Size = new System.Drawing.Size(100, 23);
            this.txtbxCell.TabIndex = 8;
            // 
            // txtbxCellPicker
            // 
            this.txtbxCellPicker.Location = new System.Drawing.Point(127, 135);
            this.txtbxCellPicker.Name = "txtbxCellPicker";
            this.txtbxCellPicker.PlaceholderText = "A1";
            this.txtbxCellPicker.Size = new System.Drawing.Size(100, 23);
            this.txtbxCellPicker.TabIndex = 9;
            // 
            // txtbxWorksheet
            // 
            this.txtbxWorksheet.Location = new System.Drawing.Point(78, 74);
            this.txtbxWorksheet.Name = "txtbxWorksheet";
            this.txtbxWorksheet.Size = new System.Drawing.Size(181, 23);
            this.txtbxWorksheet.TabIndex = 10;
            // 
            // btnReadCell
            // 
            this.btnReadCell.Location = new System.Drawing.Point(12, 134);
            this.btnReadCell.Name = "btnReadCell";
            this.btnReadCell.Size = new System.Drawing.Size(109, 23);
            this.btnReadCell.TabIndex = 11;
            this.btnReadCell.Text = "Read single cell";
            this.btnReadCell.UseVisualStyleBackColor = true;
            this.btnReadCell.Click += new System.EventHandler(this.btnReadCell_Click);
            // 
            // btnReadCellRange
            // 
            this.btnReadCellRange.Location = new System.Drawing.Point(12, 163);
            this.btnReadCellRange.Name = "btnReadCellRange";
            this.btnReadCellRange.Size = new System.Drawing.Size(109, 23);
            this.btnReadCellRange.TabIndex = 12;
            this.btnReadCellRange.Text = "Read cell range";
            this.btnReadCellRange.UseVisualStyleBackColor = true;
            this.btnReadCellRange.Click += new System.EventHandler(this.btnReadCellRange_Click);
            // 
            // txtbxCellRange
            // 
            this.txtbxCellRange.Location = new System.Drawing.Point(127, 164);
            this.txtbxCellRange.Name = "txtbxCellRange";
            this.txtbxCellRange.PlaceholderText = "A1:E7";
            this.txtbxCellRange.Size = new System.Drawing.Size(100, 23);
            this.txtbxCellRange.TabIndex = 13;
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.panel1.Controls.Add(this.btnManipulateCell);
            this.panel1.Controls.Add(this.lblCellManipulationNewValue);
            this.panel1.Controls.Add(this.lblCellManipulationOldValue);
            this.panel1.Controls.Add(this.txtbxCellManipulationCellPicker);
            this.panel1.Controls.Add(this.txtbxCellManipulationNewValue);
            this.panel1.Controls.Add(this.txtbxCellManipulationOldValue);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Location = new System.Drawing.Point(821, 34);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(260, 175);
            this.panel1.TabIndex = 14;
            // 
            // btnManipulateCell
            // 
            this.btnManipulateCell.Location = new System.Drawing.Point(16, 136);
            this.btnManipulateCell.Name = "btnManipulateCell";
            this.btnManipulateCell.Size = new System.Drawing.Size(232, 23);
            this.btnManipulateCell.TabIndex = 6;
            this.btnManipulateCell.Text = "Zelle manipulieren";
            this.btnManipulateCell.UseVisualStyleBackColor = true;
            this.btnManipulateCell.Click += new System.EventHandler(this.btnManipulateCell_Click);
            // 
            // lblCellManipulationNewValue
            // 
            this.lblCellManipulationNewValue.AutoSize = true;
            this.lblCellManipulationNewValue.Location = new System.Drawing.Point(9, 79);
            this.lblCellManipulationNewValue.Name = "lblCellManipulationNewValue";
            this.lblCellManipulationNewValue.Size = new System.Drawing.Size(68, 15);
            this.lblCellManipulationNewValue.TabIndex = 5;
            this.lblCellManipulationNewValue.Text = "neuer Wert:";
            // 
            // lblCellManipulationOldValue
            // 
            this.lblCellManipulationOldValue.AutoSize = true;
            this.lblCellManipulationOldValue.Location = new System.Drawing.Point(16, 50);
            this.lblCellManipulationOldValue.Name = "lblCellManipulationOldValue";
            this.lblCellManipulationOldValue.Size = new System.Drawing.Size(61, 15);
            this.lblCellManipulationOldValue.TabIndex = 4;
            this.lblCellManipulationOldValue.Text = "alter Wert:";
            // 
            // txtbxCellManipulationCellPicker
            // 
            this.txtbxCellManipulationCellPicker.Location = new System.Drawing.Point(83, 18);
            this.txtbxCellManipulationCellPicker.Name = "txtbxCellManipulationCellPicker";
            this.txtbxCellManipulationCellPicker.PlaceholderText = "C3";
            this.txtbxCellManipulationCellPicker.Size = new System.Drawing.Size(100, 23);
            this.txtbxCellManipulationCellPicker.TabIndex = 3;
            // 
            // txtbxCellManipulationNewValue
            // 
            this.txtbxCellManipulationNewValue.Location = new System.Drawing.Point(83, 76);
            this.txtbxCellManipulationNewValue.Name = "txtbxCellManipulationNewValue";
            this.txtbxCellManipulationNewValue.PlaceholderText = "neuer Wert";
            this.txtbxCellManipulationNewValue.Size = new System.Drawing.Size(100, 23);
            this.txtbxCellManipulationNewValue.TabIndex = 2;
            // 
            // txtbxCellManipulationOldValue
            // 
            this.txtbxCellManipulationOldValue.Location = new System.Drawing.Point(83, 47);
            this.txtbxCellManipulationOldValue.Name = "txtbxCellManipulationOldValue";
            this.txtbxCellManipulationOldValue.PlaceholderText = "alter Wert";
            this.txtbxCellManipulationOldValue.Size = new System.Drawing.Size(100, 23);
            this.txtbxCellManipulationOldValue.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(45, 21);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(32, 15);
            this.label1.TabIndex = 0;
            this.label1.Text = "Zelle";
            // 
            // lblManipulateCell
            // 
            this.lblManipulateCell.AutoSize = true;
            this.lblManipulateCell.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.lblManipulateCell.Location = new System.Drawing.Point(898, 16);
            this.lblManipulateCell.Name = "lblManipulateCell";
            this.lblManipulateCell.Size = new System.Drawing.Size(110, 15);
            this.lblManipulateCell.TabIndex = 15;
            this.lblManipulateCell.Text = "Zelle manipulieren";
            // 
            // btnSaveXlsx
            // 
            this.btnSaveXlsx.Location = new System.Drawing.Point(93, 12);
            this.btnSaveXlsx.Name = "btnSaveXlsx";
            this.btnSaveXlsx.Size = new System.Drawing.Size(75, 23);
            this.btnSaveXlsx.TabIndex = 16;
            this.btnSaveXlsx.Text = "Save .xlsx";
            this.btnSaveXlsx.UseVisualStyleBackColor = true;
            this.btnSaveXlsx.Click += new System.EventHandler(this.btnSaveXlsx_Click);
            // 
            // btnReadCsv
            // 
            this.btnReadCsv.Location = new System.Drawing.Point(12, 213);
            this.btnReadCsv.Name = "btnReadCsv";
            this.btnReadCsv.Size = new System.Drawing.Size(75, 23);
            this.btnReadCsv.TabIndex = 17;
            this.btnReadCsv.Text = "Read .csv";
            this.btnReadCsv.UseVisualStyleBackColor = true;
            this.btnReadCsv.Click += new System.EventHandler(this.btnReadCsv_Click);
            // 
            // lblCsvFile
            // 
            this.lblCsvFile.AutoSize = true;
            this.lblCsvFile.Location = new System.Drawing.Point(15, 245);
            this.lblCsvFile.Name = "lblCsvFile";
            this.lblCsvFile.Size = new System.Drawing.Size(57, 15);
            this.lblCsvFile.TabIndex = 18;
            this.lblCsvFile.Text = ".csv Datei";
            // 
            // txtbxCsvFile
            // 
            this.txtbxCsvFile.Location = new System.Drawing.Point(78, 242);
            this.txtbxCsvFile.Name = "txtbxCsvFile";
            this.txtbxCsvFile.Size = new System.Drawing.Size(401, 23);
            this.txtbxCsvFile.TabIndex = 19;
            // 
            // txtbxCsvFileView
            // 
            this.txtbxCsvFileView.Location = new System.Drawing.Point(487, 489);
            this.txtbxCsvFileView.Multiline = true;
            this.txtbxCsvFileView.Name = "txtbxCsvFileView";
            this.txtbxCsvFileView.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtbxCsvFileView.Size = new System.Drawing.Size(477, 210);
            this.txtbxCsvFileView.TabIndex = 20;
            // 
            // lblCsvFileView
            // 
            this.lblCsvFileView.AutoSize = true;
            this.lblCsvFileView.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.lblCsvFileView.Location = new System.Drawing.Point(696, 471);
            this.lblCsvFileView.Name = "lblCsvFileView";
            this.lblCsvFileView.Size = new System.Drawing.Size(79, 15);
            this.lblCsvFileView.TabIndex = 21;
            this.lblCsvFileView.Text = ".csv file view";
            // 
            // btnMergeFiles
            // 
            this.btnMergeFiles.Location = new System.Drawing.Point(12, 318);
            this.btnMergeFiles.Name = "btnMergeFiles";
            this.btnMergeFiles.Size = new System.Drawing.Size(75, 23);
            this.btnMergeFiles.TabIndex = 22;
            this.btnMergeFiles.Text = "Merge files";
            this.btnMergeFiles.UseVisualStyleBackColor = true;
            this.btnMergeFiles.Click += new System.EventHandler(this.btnMergeFiles_Click);
            // 
            // btnConvCsvXlsx
            // 
            this.btnConvCsvXlsx.Location = new System.Drawing.Point(127, 318);
            this.btnConvCsvXlsx.Name = "btnConvCsvXlsx";
            this.btnConvCsvXlsx.Size = new System.Drawing.Size(163, 23);
            this.btnConvCsvXlsx.TabIndex = 23;
            this.btnConvCsvXlsx.Text = "Convert .csv to .xlsx";
            this.btnConvCsvXlsx.UseVisualStyleBackColor = true;
            this.btnConvCsvXlsx.Click += new System.EventHandler(this.btnConvCsvXlsx_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1094, 711);
            this.Controls.Add(this.btnConvCsvXlsx);
            this.Controls.Add(this.btnMergeFiles);
            this.Controls.Add(this.lblCsvFileView);
            this.Controls.Add(this.txtbxCsvFileView);
            this.Controls.Add(this.txtbxCsvFile);
            this.Controls.Add(this.lblCsvFile);
            this.Controls.Add(this.btnReadCsv);
            this.Controls.Add(this.btnSaveXlsx);
            this.Controls.Add(this.lblManipulateCell);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.txtbxCellRange);
            this.Controls.Add(this.btnReadCellRange);
            this.Controls.Add(this.btnReadCell);
            this.Controls.Add(this.txtbxWorksheet);
            this.Controls.Add(this.txtbxCellPicker);
            this.Controls.Add(this.txtbxCell);
            this.Controls.Add(this.lblCell);
            this.Controls.Add(this.lblWorksheet);
            this.Controls.Add(this.txtbxXlsxFile);
            this.Controls.Add(this.lblXlsxFile);
            this.Controls.Add(this.txtbxXlsx);
            this.Controls.Add(this.lblXlsxFileView);
            this.Controls.Add(this.btnReadXlsx);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private Button btnReadXlsx;
        private Label lblXlsxFileView;
        private TextBox txtbxXlsx;
        private Label lblXlsxFile;
        private TextBox txtbxXlsxFile;
        private Label lblWorksheet;
        private Label lblCell;
        private TextBox txtbxCell;
        private TextBox txtbxCellPicker;
        private TextBox txtbxWorksheet;
        private Button btnReadCell;
        private Button btnReadCellRange;
        private TextBox txtbxCellRange;
        private Panel panel1;
        private Label lblManipulateCell;
        private Label lblCellManipulationNewValue;
        private Label lblCellManipulationOldValue;
        private TextBox txtbxCellManipulationCellPicker;
        private TextBox txtbxCellManipulationNewValue;
        private TextBox txtbxCellManipulationOldValue;
        private Label label1;
        private Button btnManipulateCell;
        private Button btnSaveXlsx;
        private Button btnReadCsv;
        private Label lblCsvFile;
        private TextBox txtbxCsvFile;
        private TextBox txtbxCsvFileView;
        private Label lblCsvFileView;
        private Button btnMergeFiles;
        private Button btnConvCsvXlsx;
    }
}