using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Xml;
using System.IO;
using System.Diagnostics;
using System.Reflection;

namespace Template
{
    public static class TemplateBindingUtility
    {
        /// <summary>
        /// Convert the generic list of items into a data table and return the caller
        /// </summary>
        /// <typeparam name="T">Type</typeparam>
        /// <param name="items">generic list</param>
        /// <returns>datatable</returns>
        public static DataTable ConvertListToDataTable<T>(List<T> items)

        {

            DataTable dataTable = new DataTable(typeof(T).Name);
            PropertyInfo[] Props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (PropertyInfo prop in Props)
            {
                dataTable.Columns.Add(prop.Name);
            }

            foreach (T item in items)
            {
                var values = new object[Props.Length];
                for (int i = 0; i < Props.Length; i++)
                {
                    values[i] = Props[i].GetValue(item, null);
                }
                dataTable.Rows.Add(values);
            }
            return dataTable;
        }

        /// <summary>
        /// Create the excel from the byte array/blob to the desired output path
        /// </summary>
        /// <param name="inputStream">byte array from the template</param>
        /// <param name="outputFullPath">file path+file name+ file extention</param>
        /// <returns>true if file created, false if action fails</returns>
        public static bool CreateExcel(byte[] inputStream,string outputFullPath)
        {
            try
            { 
                var fetchExcel = System.IO.File.Create(outputFullPath);
                fetchExcel.Write(inputStream, 0, inputStream.Length);
                fetchExcel.Close();
                return true;
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
        }

        /// <summary>
        /// Binds the data table input to the desired excel template in the process of generating reports
        /// </summary>
        /// <param name="dataTable">Input data table</param>
        /// <param name="excelFilename">full file path</param>
        /// <param name="sheetName">sheet where the template has to be added</param>
        /// <param name="minColumn">start column name in template like A, B, .. AG etc</param>
        /// <param name="maxcolumn">end column name in template like A, B, .. AG etc</param>
        /// <param name="minRow">start row in template like 1,2,...</param>
        /// <param name="MaxRow">end row row in template like 1,2,...</param>
        /// <returns></returns>
        public static bool BindToTemplate(DataTable dataTable, string excelFilename, string sheetName, string minColumn, string maxcolumn, uint minRow, uint MaxRow)
        {
            try
            {
                var dsRow = dataTable.Rows.OfType<DataRow>();
                uint rowIndex = minRow;
                foreach (var row in dsRow)
                {
                    char column = minColumn[0];
                    for (int columnIndex = 0; columnIndex <= ExcelColumnNameToNumber(maxcolumn) - ExcelColumnNameToNumber(minColumn); columnIndex++)
                    {
                        UpdateCell(excelFilename, sheetName, row.ItemArray[columnIndex].ToString(), rowIndex, column.ToString());
                        column++;
                    }
                    rowIndex++;
                }
                return true;
            }
            catch(Exception ex)
            {
                return false;

            }
        }
        

        /// <summary>
        /// converts the excel column name like A,B, AG etc to 1,2,33 etc
        /// </summary>
        /// <param name="columnName">column name in excel like A, B, .. AG etc</param>
        /// <returns>integer equivalent for the column</returns>
        private static int ExcelColumnNameToNumber(string columnName)
        {
            if (string.IsNullOrEmpty(columnName)) throw new ArgumentNullException("invalid column Name");
            columnName = columnName.ToUpperInvariant();
            int sum = 0;
            for (int i = 0; i < columnName.Length; i++)
            {
                sum *= 26;
                sum += (columnName[i] - 'A' + 1);
            }
            return sum;
        }

       
        /// <summary>
        /// Updates the cells based on the properties of the excel sheet
        /// </summary>
        /// <param name="docName">name of the excel</param>
        /// <param name="sheetName">name of the sheet</param>
        /// <param name="text">value to be changed</param>
        /// <param name="rowIndex">row index of the cell to be changed</param>
        /// <param name="columnName">column name of the cell to be changed</param>
        public static void UpdateCell(string docName, string sheetName,  string text,uint rowIndex, string columnName)
        {
            try
            {
                using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(docName, true))
                {
                    WorksheetPart worksheetPart = GetWorksheetPartByName(spreadSheet, sheetName);
                    if (worksheetPart != null)
                    {
                        Cell cell = GetCell(worksheetPart.Worksheet,columnName, rowIndex);
                        cell.CellValue = new CellValue(text);
                        spreadSheet.WorkbookPart.Workbook.CalculationProperties.ForceFullCalculation = true;
                        spreadSheet.WorkbookPart.Workbook.CalculationProperties.FullCalculationOnLoad = true;
                        worksheetPart.Worksheet.Save();
                        
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);

            }

        }

        /// <summary>
        /// Fetches the worksheet where we need to perform excel manipulation
        /// </summary>
        /// <param name="document">spreadsheet document</param>
        /// <param name="sheetName">sheet name</param>
        /// <returns></returns>
        private static WorksheetPart GetWorksheetPartByName(SpreadsheetDocument document,string sheetName)
        {
            IEnumerable<Sheet> sheets =document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == sheetName);
            if (sheets.Count() == 0)
            {
                return null;
            }
            string relationshipId = sheets.First().Id.Value;
            WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(relationshipId);
            return worksheetPart;

        }

        // Given a worksheet, a column name, and a row index, 
        // gets the cell at the specified column and 
        /// <summary>
        /// Given a worksheet, a column name, and a row index,gets the cell at the specified column and row index
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="columnName"></param>
        /// <param name="rowIndex"></param>
        /// <returns>Cell which is requested for</returns>
        private static Cell GetCell(Worksheet worksheet,string columnName, uint rowIndex)
        {
            Row row = GetRow(worksheet, rowIndex);
            if (row == null)
            {
                return null;
            }

            return row.Elements<Cell>().Where(c => string.Compare
                   (c.CellReference.Value, columnName +
                   rowIndex, true) == 0).First();
        }
        
        /// <summary>
        /// Given a worksheet and a row index, return the row.
        /// </summary>
        /// <param name="worksheet">worksheet</param>
        /// <param name="rowIndex">row index</param>
        /// <returns>Row that is requested for</returns>
        private static Row GetRow(Worksheet worksheet, uint rowIndex)
        {
            return worksheet.GetFirstChild<SheetData>().
              Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
        }
    }
}
