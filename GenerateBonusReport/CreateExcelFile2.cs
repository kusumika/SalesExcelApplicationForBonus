using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Data;
using System.Reflection;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System.Text.RegularExpressions;

namespace GenerateBonusReport
{
    //New changes to add fixed column width
    public class CreateExcelFile2
    {
        const int DATE_FORMAT_ID = 1;

        public static bool CreateExcelDocument<T>(List<T> list, string xlsxFilePath)
        {
            DataSet ds = new DataSet();
            ds.Tables.Add(ListToDataTable(list));

            return CreateExcelDocument(ds, xlsxFilePath);
        }

        #region HELPER_FUNCTIONS
        //  This function is adapated from: http://www.codeguru.com/forum/showthread.php?t=450171
        //  My thanks to Carl Quirion, for making it "nullable-friendly".
        public static DataTable ListToDataTable<T>(List<T> list)
        {
            DataTable dt = new DataTable();

            foreach (PropertyInfo info in typeof(T).GetProperties())
            {
                dt.Columns.Add(new DataColumn(info.Name, GetNullableType(info.PropertyType)));
            }
            foreach (T t in list)
            {
                DataRow row = dt.NewRow();
                foreach (PropertyInfo info in typeof(T).GetProperties())
                {
                    if (!IsNullableType(info.PropertyType))
                        row[info.Name] = info.GetValue(t, null);
                    else
                        row[info.Name] = (info.GetValue(t, null) ?? DBNull.Value);
                }
                dt.Rows.Add(row);
            }
            return dt;
        }

        private static Type GetNullableType(Type t)
        {
            Type returnType = t;
            if (t.IsGenericType && t.GetGenericTypeDefinition().Equals(typeof(Nullable<>)))
            {
                returnType = Nullable.GetUnderlyingType(t);
            }
            return returnType;
        }

        private static bool IsNullableType(Type type)
        {
            return (type == typeof(string) ||
                    type.IsArray ||
                    (type.IsGenericType &&
                     type.GetGenericTypeDefinition().Equals(typeof(Nullable<>))));
        }

        public static bool CreateExcelDocument(DataTable dt, string xlsxFilePath)
        {
            DataSet ds = new DataSet();
            ds.Tables.Add(dt);
            bool result = CreateExcelDocument(ds, xlsxFilePath);
            ds.Tables.Remove(dt);
            return result;
        }
        #endregion


        /// <summary>
        /// Create an Excel file, and write it to a file.
        /// </summary>
        /// <param name="ds">DataSet containing the data to be written to the Excel.</param>
        /// <param name="excelFilename">Name of file to be written.</param>
        /// <returns>True if successful, false if something went wrong.</returns>

        public static bool CreateExcelDocument(DataSet ds, string excelFilename)
        {
            try
            {
                using (SpreadsheetDocument document = SpreadsheetDocument.Create(excelFilename, SpreadsheetDocumentType.Workbook))
                {
                    WriteExcelFile(ds, document);
                }
                   return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        private static void WriteExcelFile(DataSet ds, SpreadsheetDocument spreadsheet)
        {
            //  Create the Excel file contents.  This function is used when creating an Excel file either writing 
            //  to a file, or writing to a MemoryStream.
            spreadsheet.AddWorkbookPart();
            spreadsheet.WorkbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();

            //  to prevent crashes in Excel 2010
            spreadsheet.WorkbookPart.Workbook.Append(new BookViews(new WorkbookView()));

            //  If we don't add a "WorkbookStylesPart", OLEDB will refuse to connect to this .xlsx file !
            WorkbookStylesPart workbookStylesPart = spreadsheet.WorkbookPart.AddNewPart<WorkbookStylesPart>("rIdStyles");
            Stylesheet stylesheet = new Stylesheet();
            workbookStylesPart.Stylesheet = stylesheet;

            
            //  Loop through each of the DataTables in our DataSet, and create a new Excel Worksheet for each.
            uint worksheetNumber = 1;

            foreach (DataTable dt in ds.Tables)
            {
                //  For each worksheet you want to create
                string workSheetID = "rId" + worksheetNumber.ToString();

                //KD Change
                //bool p = false;

                //if(p)
                //{
                    
                //}

                string worksheetName = dt.TableName;

                WorksheetPart newWorksheetPart = spreadsheet.WorkbookPart.AddNewPart<WorksheetPart>();
                newWorksheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet();

                // create sheet data
                newWorksheetPart.Worksheet.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.SheetData());

                // save worksheet
                WriteDataTableToExcelWorksheet(dt, newWorksheetPart);

                // To set column width in the worksheet. Created by Apt
                int numberOfColumns = dt.Columns.Count;
                for (int colInx = 1; colInx <= numberOfColumns; colInx++)
                {
                    SetColumnWidth(newWorksheetPart.Worksheet, colInx, 30);
                }
                //End
                newWorksheetPart.Worksheet.Save();

                // create the worksheet to workbook relation
                if (worksheetNumber == 1)
                    spreadsheet.WorkbookPart.Workbook.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Sheets());

                spreadsheet.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>().AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Sheet()
                {
                    Id = spreadsheet.WorkbookPart.GetIdOfPart(newWorksheetPart),
                    SheetId = (uint)worksheetNumber,
                    Name = worksheetName.Replace(':', '_').Replace('/', '_')///replacing special characters
                });
                worksheetNumber++;
            }

            spreadsheet.WorkbookPart.Workbook.Save();
        }

        private static void WriteDataTableToExcelWorksheet(DataTable dt, WorksheetPart worksheetPart)
        {
            var worksheet = worksheetPart.Worksheet;
            var sheetData = worksheet.GetFirstChild<SheetData>();

            string cellValue = "";

            //  Create a Header Row in our Excel file, containing one header for each Column of data in our DataTable.
            //
            //  We'll also create an array, showing which type each column of data is (Text or Numeric), so when we come to write the actual
            //  cells of data, we'll know if to write Text values or Numeric cell values.
            int numberOfColumns = dt.Columns.Count;
            bool[] IsNumericColumn = new bool[numberOfColumns];

            string[] excelColumnNames = new string[numberOfColumns];
            for (int n = 0; n < numberOfColumns; n++)
                excelColumnNames[n] = GetExcelColumnName(n);

            //
            //  Create the Header row in our Excel Worksheet
            //
            uint rowIndex = 1;

            var headerRow = new Row { RowIndex = rowIndex };  // add a row at the top of spreadsheet
            sheetData.Append(headerRow);
            for (int colInx = 0; colInx < numberOfColumns; colInx++)
            {
                DataColumn col = dt.Columns[colInx];
                //Created by Apt to fix column width
                uint columnIndex = Convert.ToUInt32(colInx + 1); //to add column in the excel sheet(this is excel column index)
                AppendHeaderCell(excelColumnNames[colInx] + "1", col.ColumnName, headerRow, columnIndex, worksheet);
                //end
                IsNumericColumn[colInx] = (col.ColumnName == "Bonusgrupp belopp") || (col.DataType.FullName == "System.Decimal") || (col.DataType.FullName == "System.Int32") || (col.DataType.FullName == "System.Double") || (col.DataType.FullName == "System.Single");
         
                //IsNumericColumn[colInx] = (col.DataType.FullName == "System.Decimal") || (col.DataType.FullName == "System.Int32");
            }

            //
            //  Now, step through each row of data in our DataTable...
            //
            double cellNumericValue = 0;
            foreach (DataRow dr in dt.Rows)
            {
                // ...create a new row, and append a set of this row's data to it.
                ++rowIndex;
                var newExcelRow = new Row { RowIndex = rowIndex };  // add a row at the top of spreadsheet
                sheetData.Append(newExcelRow);

                for (int colInx = 0; colInx < numberOfColumns; colInx++)
                {
                    cellValue = dr.ItemArray[colInx].ToString();

                    // Create cell with data
                    if (IsNumericColumn[colInx])
                    {
                        //  For numeric cells, make sure our input data IS a number, then write it out to the Excel file.
                        //  If this numeric value is NULL, then don't write anything to the Excel file.
                        cellNumericValue = 0;

                        cellValue = Regex.Replace(cellValue, @"[^0-9-]", "");

                        if (double.TryParse(cellValue, out cellNumericValue))
                        {
                            cellValue = cellNumericValue.ToString();
                            AppendNumericCell(excelColumnNames[colInx] + rowIndex.ToString(), cellValue, newExcelRow);
                        }
                        else 
                        { }
                    }
                    else
                    {
                        if (colInx == 2 || colInx == 3)
                        {
                            cellValue = cellValue.Substring(0, 9);
                        }
                        //  For text cells, just write the input data straight out to the Excel file.
                        AppendTextCell(excelColumnNames[colInx] + rowIndex.ToString(), cellValue, newExcelRow);
                    }
                }
            }

            // Created by APT to add formula row 17-09-2015
            //Added additional cell for sum of certain columns
            rowIndex = rowIndex + 1;
            //var newFormulaRow = new Row { RowIndex = rowIndex };  // add a row at the top of spreadsheet
            //sheetData.Append(newFormulaRow);
            //foreach (DataColumn dc in dt.Columns)
            //{
            //    string formula = "=SUM(" + excelColumnNames[dc.Ordinal] + "2:" + excelColumnNames[dc.Ordinal] + (rowIndex - 1).ToString() + ")";
            //    switch (dc.ColumnName)
            //    {
            //        case "Medlemsort":
            //            AppendTextCell(excelColumnNames[dc.Ordinal] + rowIndex.ToString(), "Totalt", newFormulaRow);
            //            break;
            //        case "Bonusgrupp inköpsbelopp":
            //            AppendFormulaCell(excelColumnNames[dc.Ordinal] + rowIndex.ToString(), formula, newFormulaRow);
            //            break;
            //        case "Totalt inköpsbelopp":
            //            AppendFormulaCell(excelColumnNames[dc.Ordinal] + rowIndex.ToString(), formula, newFormulaRow);
            //            break;
            //        case "Bonus i kr":
            //            AppendFormulaCell(excelColumnNames[dc.Ordinal] + rowIndex.ToString(), formula, newFormulaRow);
            //            break;
            //        case "Total bonus i kr":
            //            AppendFormulaCell(excelColumnNames[dc.Ordinal] + rowIndex.ToString(), formula, newFormulaRow);
            //            break;
            //    }
            //}
        }

        //This function is created by Apt to append header cell with corresponding column
        private static void AppendHeaderCell(string cellReference, string cellStringValue, Row excelRow, uint columnIndex, DocumentFormat.OpenXml.Spreadsheet.Worksheet worksheet)
        {
            DocumentFormat.OpenXml.Spreadsheet.Columns columns;
            DocumentFormat.OpenXml.Spreadsheet.Column previousColumn = null;
            //  Add a new Excel Cell to our Row 
            Cell cell = new Cell() { CellReference = cellReference, DataType = CellValues.String };
            CellValue cellValue = new CellValue();
            cellValue.Text = cellStringValue;
            cell.Append(cellValue);
            excelRow.Append(cell);

            columns = worksheet.Elements<DocumentFormat.OpenXml.Spreadsheet.Columns>().FirstOrDefault();
            // Check if the column collection exists
            if (columns == null)
            {
                columns = worksheet.InsertAt(new DocumentFormat.OpenXml.Spreadsheet.Columns(), 0);
            }
            // Check if the column exists
            if (columns.Elements<DocumentFormat.OpenXml.Spreadsheet.Column>().Where(item => item.Min == columnIndex).Count() == 0)
            {
                // Find the previous existing column in the columns
                for (uint counter = columnIndex - 1; counter > 0; counter--)
                {
                    previousColumn = columns.Elements<DocumentFormat.OpenXml.Spreadsheet.Column>().Where(item => item.Min == counter).FirstOrDefault();
                    if (previousColumn != null)
                    {
                        break;
                    }
                }
                columns.InsertAfter(
                   new DocumentFormat.OpenXml.Spreadsheet.Column()
                   {
                       Min = columnIndex,
                       Max = columnIndex,
                       CustomWidth = true,
                       Width = 9
                   }, previousColumn);
            }
        }
        //Changes by APT
        private static void AppendTextCell(string cellReference, string cellStringValue, Row excelRow)
        {
            //  Add a new Excel Cell to our Row 
            Cell cell = new Cell() { CellReference = cellReference, DataType = CellValues.String };
            CellValue cellValue = new CellValue();
            cellValue.Text = cellStringValue;
            cell.Append(cellValue);
            excelRow.Append(cell);
        }

        private static void AppendNumericCell(string cellReference, string cellStringValue, Row excelRow)
        {
            //  Add a new Excel Cell to our Row 
            Cell cell = new Cell() { CellReference = cellReference, DataType = CellValues.Number };
            CellValue cellValue = new CellValue();
            cellValue.Text = cellStringValue.Replace(",", ".");//cellStringValue;
            cell.Append(cellValue);
            excelRow.Append(cell);
        }
        //End

        //End
        //Created by Apt 17-09-2015
        private static void AppendFormulaCell(string cellReference, string cellStringValue, Row excelRow)
        {
            //  Add a new Excel Cell to our Row 
            Cell cell = new Cell() { CellReference = cellReference, DataType = CellValues.Number };
            CellFormula cellFormula = new CellFormula() { CalculateCell = true, Text = cellStringValue };
            cell.Append(cellFormula);
            excelRow.Append(cell);
        }
        //End

        private static string GetExcelColumnName(int columnIndex)
        {
            //  Convert a zero-based column index into an Excel column reference  (A, B, C.. Y, Y, AA, AB, AC... AY, AZ, B1, B2..)
            //
            //  eg  GetExcelColumnName(0) should return "A"
            //      GetExcelColumnName(1) should return "B"
            //      GetExcelColumnName(25) should return "Z"
            //      GetExcelColumnName(26) should return "AA"
            //      GetExcelColumnName(27) should return "AB"
            //      ..etc..
            if (columnIndex < 26)
                return ((char)('A' + columnIndex)).ToString();

            char firstChar = (char)('A' + (columnIndex / 26) - 1);
            char secondChar = (char)('A' + (columnIndex % 26));

            return string.Format("{0}{1}", firstChar, secondChar);
        }
        /// <summary>
        /// Sets the column width
        /// </summary>
        /// <param name="worksheet">Worksheet to use</param>
        /// <param name="columnIndex">Index of the column</param>
        /// <param name="width">Width to set</param>
        /// <returns>True if succesful</returns>
        public static bool SetColumnWidth(DocumentFormat.OpenXml.Spreadsheet.Worksheet worksheet, int columnIndex, int width)
        {
            DocumentFormat.OpenXml.Spreadsheet.Columns columns;
            DocumentFormat.OpenXml.Spreadsheet.Column column;

            // Get the column collection exists
            columns = worksheet.Elements<DocumentFormat.OpenXml.Spreadsheet.Columns>().FirstOrDefault();
            if (columns == null)
            {
                return false;
            }
            // Get the column
            column = columns.Elements<DocumentFormat.OpenXml.Spreadsheet.Column>().Where(item => item.Min == columnIndex).FirstOrDefault();
            if (column == null)
            {
                return false;
            }
            column.Width = width;
            column.CustomWidth = true;
            worksheet.Save();

            return true;
        }
    }
}