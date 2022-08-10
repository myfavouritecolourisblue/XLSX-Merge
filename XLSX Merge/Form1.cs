using ClosedXML.Excel;
using Microsoft.VisualBasic.FileIO;

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

            csvToWorksheet(filepath, newSheet);

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
            #region CSV file import
            // read csv into temporary worksheet
            string filepath = "C:\\temp\\quelle.csv";
            if (!File.Exists(filepath)) {
                MessageBox.Show("Error: CSV file doesn't exist. Aborting operation.");
                return;
            }
            
            // Create temporary Excel workbook
            XLWorkbook tempCsvWb = new XLWorkbook();
            // Create temporary Excel sheet
            IXLWorksheet tempCsvWs = tempCsvWb.AddWorksheet("csv-import");

            // Import CSV data into worksheet
            try {
                csvToWorksheet(filepath, tempCsvWs);
            } catch (Exception ex) {
                MessageBox.Show("An error occured while importing the csv file. Error message:\r\n\r\n" + ex.Message);
                return;
            }

            // get the mapping of header-name-string to header-position-int
            Dictionary<string, int> csvHeaderXPositionKvp = new Dictionary<string, int>();
            IXLRow firstRow = tempCsvWs.FirstRow();

            foreach (var c in firstRow.CellsUsed())
                csvHeaderXPositionKvp.Add(c.GetString(), c.Address.ColumnNumber);

            string indexHeader = txtbxMergeHeader.Text;
            int indexHeaderNr = csvHeaderXPositionKvp[indexHeader];

            // sort data by the index
            //tempCsvWs.Sort(indexHeader);
            #endregion


            // open xlsx file, open workbook, open worksheet (maybe as a stream instead of a file)
            // TODO: open as a stream
            string filepathDest = "C:\\temp\\quelle.xlsx";
            if (!File.Exists(filepathDest)) {
                MessageBox.Show("Error: Excel file does not exist. Aborting operation.");
                return;
            }

            IXLWorkbook destinationWb;
            try {
                destinationWb = new XLWorkbook(filepathDest);
            } catch (System.IO.IOException ex) {
                MessageBox.Show("Error: Opening the Excel file failed. Error message:\r\n\r\n" + ex.Message);
                return;
            }
            IXLWorksheet destinationWs = destinationWb.Worksheet(1);

            // this will be our Y-coordinate, the row number in which the headers are contained in the existing Excel file
            int? destHeaderRowNr = null;

            // sort the .csv header row for a List comparison with the .xlsx rows
            List<string> csvHeaderList = new();
            foreach (IXLCell c in firstRow.CellsUsed().ToList())
                csvHeaderList.Add(c.GetString());

            foreach (IXLRow r in destinationWs.RowsUsed())
            {
                // Get all cell values in a List and sort them
                List<string> destRowList = new();
                foreach (IXLCell c in r.CellsUsed().ToList())
                    destRowList.Add(c.GetString());

                // compare the .xlsx row with the .csv header row
                bool headerFoundInRow = true;
                foreach (string s in csvHeaderList)
                    if (!destRowList.Contains(s))
                        headerFoundInRow = false;
                if (!headerFoundInRow)
                    continue;
                
                destHeaderRowNr = r.RowNumber();

                break;
            }

            // Abort if no fitting row was found
            if (destHeaderRowNr is null) {
                MessageBox.Show("Error: Corresponding column headers of the CSV file not found in Excel file or the column headers are placed in different rows. Aborting operation.");
                return;
            }

            //  X-coordinates of each header
            Dictionary<string,int> xlsxHeaderXPositionKvp = new();

            IXLRow destinationRow = destinationWs.Row((int)destHeaderRowNr);
            // For each header in our CSVs header dictionary ...
            foreach (string s in csvHeaderXPositionKvp.Keys) {
                IXLCells c = destinationRow.Search(s); // ... search the row for cells containing the header
                xlsxHeaderXPositionKvp.Add(s, c.First().Address.ColumnNumber);   // ... and add the first found cell's column number (X-coordinate) as value to the dict
            }

            
            // Check for merging method
            if (cbMergeMethod.Text.Equals("Append"))
            {
                // Check for next empty cell in the index header column (the Y-coordinate) and increase it by 1 to get the next free cell
                int startrowOfRangeInsert = destinationWs.Column(xlsxHeaderXPositionKvp[indexHeader]).LastCellUsed().Address.RowNumber + 1;

                /* Get the number of entries in the csv header column by
                 * subtracting 1 off of the last used cell's row number. In
                 * case of an empty cell in the indexheader column in between
                 * the function counts adds the row to the range length as long
                 * as somewhere further down is a cell with a value. */
                int rangeLength = tempCsvWs.Column(indexHeaderNr).LastCellUsed().Address.RowNumber - 1;

                // Insert each presorted column (from Step 2) vertically at the first free row and the X-coordinate of the headers column
                foreach (KeyValuePair<string, int> csvKvp in csvHeaderXPositionKvp)
                {
                    IXLCell startCell = destinationWs.Cell(startrowOfRangeInsert, xlsxHeaderXPositionKvp[csvKvp.Key]);

                    // Construct the vertical range that holds the csv data
                    IXLRange dataRange = tempCsvWs.Range(2, csvKvp.Value, rangeLength + 1, csvKvp.Value);

                    // Insert the values
                    startCell.Value = dataRange;
                }
            } else if (cbMergeMethod.Text.Equals("Replace")) {
                // The insert range starts just below the header row
                int startrowOfRangeInsert = (int)destHeaderRowNr + 1;

                // Get the number of entries in the csv
                int rangeLength = tempCsvWs.Column(indexHeaderNr).CellsUsed().Count() - 1;

                // Insert each presorted column (from Step 2) vertically in the row below the header and the X-coordinate of the headers column
                // and delete all other entries for this column under the header
                foreach (KeyValuePair<string, int> csvKvp in csvHeaderXPositionKvp)
                {
                    IXLCell startCell = destinationWs.Cell(startrowOfRangeInsert, xlsxHeaderXPositionKvp[csvKvp.Key]);

                    #region Clean entries before inserting the new data
                    IXLCell removeRangeEndCell = destinationWs.Cell(
                        destinationWs.Column(xlsxHeaderXPositionKvp[csvKvp.Key]).LastCellUsed().Address.RowNumber,
                        xlsxHeaderXPositionKvp[csvKvp.Key]
                    );

                    // If the end cell has a lower row number, it means that there are no entries in this column underneath the header,
                    // so the start cell is equal to the end cell.
                    if (removeRangeEndCell.Address.RowNumber < startCell.Address.RowNumber)
                        removeRangeEndCell = startCell;

                    destinationWs.Range(startCell, removeRangeEndCell).Clear();
                    #endregion


                    #region Insert data
                    // Construct the vertical range that holds the csv data
                    IXLRange dataRange = tempCsvWs.Range(2, csvKvp.Value, rangeLength + 1, csvKvp.Value);

                    // Insert the values
                    startCell.Value = dataRange;
                    #endregion
                }
            } else { 
                return; // Abort execution
            }

            // Save workbook
            destinationWb.Save();
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

        #region Utility Funktionen und noch auszulagernde Funktionen
        ///////////////////////////////////////////////////////////////////////
        /// NOCH AUSLAGERN IN ANDERE DATEI ///
        ///////////////////////////////////////////////////////////////////////

        private static void csvToWorksheet(string filepath, IXLWorksheet newSheet)
        {
            using (TextFieldParser parser = new(filepath))
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
                    if (line == null) { break; }
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
        }
        #endregion
    }
}