using ClosedXML.Excel;
using Microsoft.VisualBasic.FileIO;
using System.Diagnostics;
using XLSX_Merge_Utils;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;

namespace XLSX_Merge
{
    public partial class Form1 : Form
    {
        #region JaHierSindDefinitionen
        // The actual Excel file
        XLWorkbook workbook;

        // The current Excel sheet
        IXLWorksheet worksheet;

        #endregion

        public Form1()
        {
            InitializeComponent();
        }

        /*
         * Loads the excel file and the sheet on first position into memory.
         */
        private void btnReadXlsx_Click(object sender, EventArgs e)
        {
            if (Environment.GetCommandLineArgs().Contains("--noui")) this.Hide();
            /*string filepath = "C:\\temp\\quelle.xlsx";
            
            // Load the Excel file into memory
            workbook = new XLWorkbook(filepath);

            // Show file path of the loaded file to the user
            txtbxXlsxFile.Text = filepath;

            // Load worksheet on position 1 into memory
            worksheet = workbook.Worksheet(1);

            // Show the worksheet name to the user
            txtbxWorksheet.Text = worksheet.ToString();*/
        }

        /*
         * Reads and displays a single cells value.
         */
        private void btnReadCell_Click(object sender, EventArgs e)
        {
            // Check if
            // 1. a cell id was given by the user
            // 2. the worksheet to be accessed is loaded into memory
            if (txtbxCellPicker.Text.Equals("") || worksheet.Equals(null)) return;

            
            txtbxXlsx.Clear();

            try
            {
                IXLCell cell = worksheet.Cell(txtbxCellPicker.Text);    // Select cell based on the id
                string cellValueString = cell.GetValue<string>();       // Extract cell value
                txtbxXlsx.Text = cellValueString;                       // Show cell value to user
            } catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /*
         * Reads and displays the values of a range of cells.
         */
        private void btnReadCellRange_Click(object sender, EventArgs e)
        {
            // Check if
            // 1. a cell range was given by the user
            // 2. the worksheet to be accessed is loaded into memory
            if (txtbxCellRange.Text.Equals("") || worksheet.Equals(null)) return;
            
            txtbxXlsx.Clear();

            // Gute Doku: https://github.com/ClosedXML/ClosedXML
            try
            {
                IXLRange range = worksheet.Range(txtbxCellRange.Text);  // Select a range of cells based on the user given range
                foreach (var cell in range.Cells())                       // Extract each cells value and display it in the textbox
                    txtbxXlsx.Text = txtbxXlsx.Text 
                                    + cell.GetString() 
                                    + "\r\n";

            } catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }

        /*
         * Manipulates a single cells value.
         */
        private void btnManipulateCell_Click(object sender, EventArgs e)
        {
            // Check if
            // 1. a cell id was given by the user
            // 2. the cells new value is available
            if (txtbxCellManipulationCellPicker.Text.Equals("") || txtbxCellManipulationNewValue.Text.Equals("")) return;

            try
            {
                // Read and display the cells old value to the user
                txtbxCellManipulationOldValue.Text = worksheet.Cell(txtbxCellManipulationCellPicker.Text).GetValue<string>();
                // Set the cells new value
                worksheet.Cell(txtbxCellManipulationCellPicker.Text).SetValue<string>(txtbxCellManipulationNewValue.Text);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /*
         * Saves the workbook as an Excel file to a directory.
         */
        private void btnSaveXlsx_Click(object sender, EventArgs e)
        {
            workbook.SaveAs("C:\\temp\\neu_" + DateTime.Now.Ticks + ".xlsx");
        }

        // Experimental. No defined use.
        private void btnReadCsv_Click(object sender, EventArgs e)
        {
            /*
             * https://stackoverflow.com/questions/2081418/parsing-csv-files-in-c-with-header
             * https://stackoverflow.com/questions/5282999/reading-csv-file-and-storing-values-into-an-array
             */
            string filepath = "C:\\temp\\quelle.csv";
            txtbxCsvFile.Text = filepath;

            using(TextFieldParser parser = new TextFieldParser(filepath))
            {
                parser.TextFieldType = FieldType.Delimited;
                parser.SetDelimiters(";");

                while (!parser.EndOfData)
                {
                    string[] fields = parser.ReadFields();
                    foreach(string field in fields)
                    {
                        var old = txtbxCsvFileView.Text;
                        txtbxCsvFileView.Clear();
                        txtbxCsvFileView.Text = old + "\r\n" + field;
                    }
                }
            }
        }

        /*
         * Loads a .csv file into memory and converts it into a .xlsx file.
         */
        private void btnConvCsvXlsx_Click(object sender, EventArgs e)
        {
            string filepath = "C:\\temp\\quelle.csv";
            txtbxCsvFile.Text = filepath;

            // Create new Excel workbook
            XLWorkbook newFile = new XLWorkbook();
            // Create new Excel sheet
            IXLWorksheet newSheet = newFile.AddWorksheet("Tabelle123");
            XlsxMergeUtils.csvToWorksheet(filepath, newSheet);

            newFile.SaveAs("C:\\temp\\neu_" + DateTime.Now.Ticks + ".xlsx");
        }

        
        // TODO: Tests, Fehlermeldungen, Dokumentation, commandline arguments, open file dialog
        /// <summary>
        /// Merges a .csv file into a .xlsx file.
        /// It uses the headers (first row values) of the .csv file as an orientation,
        /// looks for these headers in the .xlsx file and either
        /// 1) appends the .csv file's values just after the last value below the 
        /// corresponding header,
        /// OR
        /// 2) clears all values below the headers and inserts the .csv's data below
        ///     the corresponding header.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnMergeFiles_Click(object sender, EventArgs e) {
            switch (cbMergeMethod.Text) {
                case "Append":
                    XlsxMergeUtils.mergeCSVintoXLSX(txtbxCsvFile.Text, txtbxXlsxFile.Text, txtbxMergeHeader.Text, XlsxMergeUtils.MergeMethods.Append);
                    break;
                case "Replace":
                    XlsxMergeUtils.mergeCSVintoXLSX(txtbxCsvFile.Text, txtbxXlsxFile.Text, txtbxMergeHeader.Text, XlsxMergeUtils.MergeMethods.Replace);
                    break;
                default:
                    break;
            }

            txtbxXlsx.Text = "Fertig" + DateTime.Now;
        }


        #region Funktionen die ich nicht löschen kann
        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void lblMergeDataHeader_Click(object sender, EventArgs e)
        {

        }
        #endregion

        private void openXlsxFileDialog_FileOk(object sender, System.ComponentModel.CancelEventArgs e) {
        }

        private void btnSelCsvFile_Click(object sender, EventArgs e) {
            openFileDialog.InitialDirectory = "C:\\Users\\" + Environment.UserName + "\\Documents\\";
            openFileDialog.Title = "Select .csv file";
            openFileDialog.ShowDialog();
            txtbxCsvFile.Text = openFileDialog.FileName;
        }

        private void btnSelXlsxFile_Click(object sender, EventArgs e) {
            openFileDialog.InitialDirectory = "C:\\Users\\" + Environment.UserName + "\\Documents\\";
            openFileDialog.Title = "Select .xlsx file";
            openFileDialog.ShowDialog();
            
            txtbxXlsxFile.Text = openFileDialog.FileName;
        }

        private void openFileDialog_FileOk(object sender, System.ComponentModel.CancelEventArgs e) {

        }
    }
}