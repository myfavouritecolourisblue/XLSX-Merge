using ClosedXML.Excel;
using Microsoft.VisualBasic.FileIO;

namespace XLSX_Merge_Utils {
    public class XlsxMergeUtils {

        public enum MergeMethods {
            Append,
            Replace
        }

        private static Dictionary<string,int> mapHeaderToColumnNr(IXLRow headerRow) {
            Dictionary<string, int> csvHeaderXPositionKvp = new Dictionary<string, int>();

            foreach (var c in headerRow.CellsUsed())
                csvHeaderXPositionKvp.Add(c.GetString(), c.Address.ColumnNumber);

            return csvHeaderXPositionKvp;
        }

        private static XLWorkbook openXlsxFile(string xlsxFilepath, FileStream? xlsxFS) {
            // open xlsx file, open workbook, open worksheet (maybe as a stream instead of a file)
            // TODO: open as a stream

            XLWorkbook destinationWb;

            try {
                //destinationWb = new XLWorkbook(filepathDest);
                xlsxFS = new FileStream(xlsxFilepath, FileMode.Open, FileAccess.ReadWrite);
                destinationWb = new XLWorkbook(xlsxFS);
            } catch (FileNotFoundException ex) {
                if (xlsxFS != null)
                    xlsxFS.Close();

                Console.WriteLine("Error: Excel file does not exist. Aborting operation.");
                throw;
            } catch (Exception ex) {
                if (xlsxFS != null)
                    xlsxFS.Close();

                Console.WriteLine("Error: Opening the Excel file failed. Error message:\r\n\r\n" + ex.Message);
                throw;
            }
            return destinationWb;
        }

        /// <summary>
        /// Merges a .csv file into a .xlsx file. A unique key to the table, existing only once globally.
        /// </summary>
        /// <param name="indexHeader">The index, the point of orientation when deciding where and how much to insert.</param>
        public static void mergeCSVintoXLSX(string csvFilePath, string xlsxFilePath, string indexHeader, MergeMethods mergeMethod) {
            #region CSV file import
            // read csv into temporary worksheet

            // Create temporary Excel workbook
            XLWorkbook tempCsvWb = new XLWorkbook();

            // Create temporary Excel sheet
            IXLWorksheet tempCsvWs = tempCsvWb.AddWorksheet("csv-import");
            
            csvToWorksheet(csvFilePath, tempCsvWs);

            #endregion

            #region  Create mapping of header name-string to column-number
            Dictionary<string,int> csvHeaderXPositionKvp = mapHeaderToColumnNr(tempCsvWs.FirstRow());
            #endregion

            #region Open .xlsx file
            FileStream? xlsxFS = null;
            IXLWorkbook destinationWb = openXlsxFile(xlsxFilePath, xlsxFS);
            #endregion

            // TODO: Make the sheet number variable
            IXLWorksheet destinationWs = destinationWb.Worksheet(1);

            #region Search for the row in which the data headers are located and get its row number
            // Get the .csv data headers for comparison with the .xlsx headers
            #region Convert .csv header row to a List of strings
            List<string> csvHeaderList = new();
            foreach (IXLCell c in tempCsvWs.FirstRow().CellsUsed().ToList())
                csvHeaderList.Add(c.GetString());
            #endregion
            // The number of the row in which the headers are contained in the existing Excel file
            int? destHeaderRowNr = null;

            foreach (IXLRow r in destinationWs.RowsUsed()) {
                #region Convert .xlsx header row to a List of strings
                // Get all cell values as strings in a List
                List<string> destRowList = new();
                foreach (IXLCell c in r.CellsUsed().ToList())
                    destRowList.Add(c.GetString());
                #endregion
                #region Compare both header rows
                // compare the .xlsx row with the .csv header row
                bool headerFoundInRow = true;
                foreach (string s in csvHeaderList)
                    if (!destRowList.Contains(s))
                        headerFoundInRow = false;
                #endregion
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

            #endregion

            #region Build a dict with the .xlsx headers paired with its respective column numbers
            // A dict with the headers name string paired with its column number
            Dictionary<string, int> xlsxHeaderXPositionKvp = new();

            IXLRow destinationRow = destinationWs.Row((int)destHeaderRowNr);
            // For each header in our CSVs header dictionary ...
            foreach (string s in csvHeaderXPositionKvp.Keys) {
                IXLCells c = destinationRow.Search(s); // ... search the row for cells containing the header
                xlsxHeaderXPositionKvp.Add(s, c.First().Address.ColumnNumber);   // ... and add the first found cell's column number as value to the dict
            }
            #endregion

            #region Determination the index header's column number in the .csv file
            if (string.IsNullOrEmpty(indexHeader)) {
                MessageBox.Show("Merge header not given. Aborting operation.");
                return;
            }
            int indexHeaderNr = csvHeaderXPositionKvp[indexHeader];
            #endregion

            #region Perform the actual merge
                #region Merge by appending
            // Check for merging method
            if (mergeMethod.Equals(MergeMethods.Append)) {
                // Check for last used cell in the indexHeader column and increase its row number 1 to get the next free cell
                int startrowOfRangeInsert = destinationWs.Column(xlsxHeaderXPositionKvp[indexHeader]).LastCellUsed().Address.RowNumber + 1;

                /* Get the number of entries in the csv header column by
                 * subtracting 1 off of the last used cell's row number. In
                 * case of an empty cell in the indexheader column in between
                 * the function counts adds the row to the range length as long
                 * as somewhere further down is a cell with a value. */
                int rangeLength = tempCsvWs.Column(indexHeaderNr).LastCellUsed().Address.RowNumber - 1;

                // Insert the csv's data vertically at the first free row in its respective column under the header
                foreach (KeyValuePair<string, int> csvKvp in csvHeaderXPositionKvp) {
                    IXLCell startCell = destinationWs.Cell(startrowOfRangeInsert, xlsxHeaderXPositionKvp[csvKvp.Key]);

                    // Construct the vertical range that holds the csv data
                    IXLRange dataRange = tempCsvWs.Range(2, csvKvp.Value, rangeLength + 1, csvKvp.Value);

                    // Insert the values
                    startCell.Value = dataRange;
                }
                #endregion
                #region Merge by replacing
            } else if (mergeMethod == MergeMethods.Replace) {
                // The insert range starts just below the header row
                int startrowOfRangeInsert = (int)destHeaderRowNr + 1;

                // Get the number of entries in the csv
                int rangeLength = tempCsvWs.Column(indexHeaderNr).CellsUsed().Count() - 1;

                /* Insert the csv's data vertically at the first free row in 
                 * its respective column under the header and delete all 
                 * other entries for this column under the header. */
                foreach (KeyValuePair<string, int> csvKvp in csvHeaderXPositionKvp) {
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
                #endregion
                #region Error: No merge method given
            } else {
                MessageBox.Show("Neither \"Append\" nor \"Replace\" was given as the merge method. Aborting operation.");
                return; // Abort execution
            }
            #endregion
            #endregion

            #region Save changes to Excel file
            // Save workbook
            destinationWb.Save();
            if (xlsxFS != null) {
                xlsxFS.Flush();
                xlsxFS.Close();
            }
            #endregion
        }





        #region Utility Funktionen

        public static void csvToWorksheet(string filepath, IXLWorksheet newSheet) {
            try {
                using (TextFieldParser parser = new(filepath)) {
                    // Tell parser that the text is delimited with a semicolon
                    parser.TextFieldType = FieldType.Delimited;
                    parser.SetDelimiters(";");

                    // Rows are counted from 1 instead of 0, that means:
                    // row == arrayindex + 1
                    int lineNumber = 1;

                    while (!parser.EndOfData) {
                        // Read the next line of the csv file
                        string? line = parser.ReadLine();
                        if (line == null) { break; } else {
                            // Split current line by its delimiter
                            string[] splittedValues = line.Split(";");
                            // The number of substrings equals the number of columns
                            int numberOfColumns = splittedValues.Length;

                            // Go through each column in the current row (line) and set the cells value
                            for (int i = 0; i < numberOfColumns; i++) {
                                int row = lineNumber;
                                int column = i + 1;

                                newSheet.Cell(row, column).SetValue(splittedValues[i]);
                            }
                        }
                        lineNumber++;
                    }
                }
            } catch (FileNotFoundException ex) {
                Console.WriteLine("Error: CSV file doesn't exist. Aborting operation.");
                throw;
            } catch {
                throw;
            }
        }

        #endregion
    }
}