using ClosedXML.Excel;
using Microsoft.VisualBasic.FileIO;

namespace XLSX_Merge
{
    public partial class Form1 : Form
    {
        // The actual Excel file
        XLWorkbook workbook;

        // The current Excel sheet
        IXLWorksheet worksheet;

        public Form1()
        {
            InitializeComponent();
        }

        /*
         * Loads the excel file and the sheet on first position into memory.
         */
        private void btnReadXlsx_Click(object sender, EventArgs e)
        {
            string filepath = "C:\\temp\\quelle.xlsx";

            // Load the Excel file into memory
            workbook = new XLWorkbook(filepath);

            // Show file path of the loaded file to the user
            txtbxXlsxFile.Text = filepath;

            // Load worksheet on position 1 into memory
            worksheet = workbook.Worksheet(1);

            // Show the worksheet name to the user
            txtbxWorksheet.Text = worksheet.ToString();
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
                foreach (var cell in range.Cells())                     // Extract each cells value and display it in the textbox
                    txtbxXlsx.Text = txtbxXlsx.Text 
                                    + cell.GetValue<string>() 
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

            using (TextFieldParser parser = new TextFieldParser(filepath))
            {
                // Tell parser that the text is delimited with a semicolon
                parser.TextFieldType = FieldType.Delimited;
                parser.SetDelimiters(";");

                // Rows a counted from 1 instead of 0, that means:
                // row == arrayindex + 1
                int lineNumber = 1;

                while (!parser.EndOfData)
                {
                    // Read the next line of the csv file
                    string? line = parser.ReadLine();
                    if(line == null) { break; }
                    else
                    {
                        // Split current line by its delimiter
                        string[] splittedValues = line.Split(";");
                        // The number of substrings equals the number of columns
                        int numberOfColumns = splittedValues.Length;

                        // Go through each column in the current row (line) and set the cells value
                        for (int i = 0; i < numberOfColumns; i++)
                        {
                            int row = lineNumber;
                            int column = i + 1;

                            newSheet.Cell(row, column).SetValue(splittedValues[i]);
                        }
                    }
                    lineNumber++;
                }
            }

            newFile.SaveAs("C:\\temp\\neu_" + DateTime.Now.Ticks + ".xlsx");
        }

        private void btnMergeFiles_Click(object sender, EventArgs e)
        {

        }
    }
}