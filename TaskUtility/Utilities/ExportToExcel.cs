using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace TaskUtility
{
    public sealed class ExportToExcelUtility

    {

        /// <summary>

        /// Convert the list to Data table.

        /// </summary>

        /// <param name="items">List</param>

        /// <returns>DataTable</returns>

        public static DataTable ToDataTable<T>(List<T> items)

        {

            DataTable dataTable = new DataTable(typeof(T).Name);

            PropertyInfo[] Props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);

            foreach (PropertyInfo prop in Props)

            {

                dataTable.Columns.Add(prop.Name);//Setting column names

            }

            foreach (T item in items)

            {

                var values = new object[Props.Length];

                for (int i = 0; i < Props.Length; i++)

                {

                    values[i] = Props[i].GetValue(item, null);//Insert values.

                }

                dataTable.Rows.Add(values);

            }

            return dataTable;

        }

        /// <summary>

        /// Call this method from the Page providing the desired information

        /// </summary>

        /// <param name="dataTable">The records to be written in excel</param>

        /// <param name="excelFilename">Name of the file</param>

        /// <param name="sheetName">Name of the sheet</param>

        /// <param name="filters">Search key and value based on which the datatable is generated</param>

        /// <param name="columnSize">column name and size</param>

        /// <returns></returns>

        public static bool CreateExcelDocument(DataTable dataTable, string excelFilename, string sheetName, Dictionary<string, string> filters, Dictionary<string, int> columnSize)

        {

            try

            {

                using (SpreadsheetDocument objExcelDoc = SpreadsheetDocument.Create(excelFilename, SpreadsheetDocumentType.Workbook))

                {

                    int cellSize;

                    WorkbookPart wbp = objExcelDoc.AddWorkbookPart();

                    WorksheetPart wsp = wbp.AddNewPart<WorksheetPart>();

                    Workbook wb = new Workbook();

                    FileVersion fv = new FileVersion();

                    fv.ApplicationName = "Microsoft Office Excel";

                    Worksheet workSheet = new Worksheet();

                    WorkbookStylesPart wbsp = wbp.AddNewPart<WorkbookStylesPart>();

                    wbsp.Stylesheet = CreateStylesheet();

                    wbsp.Stylesheet.Save();

                    Columns columns = new Columns();

                    for (int i = 1; i <= columnSize.Count(); i++)

                    {

                        columnSize.TryGetValue(columnSize.Keys.ElementAt(i - 1).ToString(), out cellSize);

                        columns.Append(CreateColumnData(Convert.ToUInt32(i), Convert.ToUInt32(i), cellSize));

                    }

                    workSheet.Append(columns);

                    SheetData sheetData = new SheetData();

                    for (UInt32 i = 2; i <= 1 + filters.Count(); i++)

                    {

                        sheetData.Append(CreateFilters(i, filters));

                    }

                    sheetData.Append(CreateColumnHeader(Convert.ToUInt32(filters.Count() + 3), columnSize));

                    UInt32 index = Convert.ToUInt32(filters.Count() + 4);

                    foreach (DataRow dr in dataTable.Rows)

                    {

                        sheetData.Append(CreateContent(index, dr, columnSize.Count()));

                        index++;

                    }

                    workSheet.Append(sheetData);

                    wsp.Worksheet = workSheet;

                    Sheets sheets = new Sheets();

                    Sheet sheet = new Sheet();

                    sheet.Name = sheetName;

                    sheet.SheetId = 1;

                    sheet.Id = wbp.GetIdOfPart(wsp);

                    sheets.Append(sheet);

                    wb.Append(fv);

                    wb.Append(sheets);

                    objExcelDoc.WorkbookPart.Workbook = wb;

                    objExcelDoc.WorkbookPart.Workbook.Save();

                    objExcelDoc.Close();

                }

            }

            catch (Exception ex)

            {

                throw;

            }

            return true;

        }

        /// <summary>

        /// Create column for storing data by passing the start column index, end column index and column width

        /// </summary>

        /// <param name="StartColumnIndex">start column index</param>

        /// <param name="EndColumnIndex">end column index</param>

        /// <param name="ColumnWidth">width of each column</param>

        /// <returns>column</returns>

        private static Column CreateColumnData(UInt32 StartColumnIndex, UInt32 EndColumnIndex, double ColumnWidth)

        {

            Column column;

            column = new Column();

            column.Min = StartColumnIndex;

            column.Max = EndColumnIndex;

            column.Width = ColumnWidth;

            column.CustomWidth = true;

            return column;

        }

        /// <summary>

        /// Writes the row to the excel by reading each datarow from the datatable

        /// </summary>

        /// <param name="index">row index</param>

        /// <param name="dr">data row</param>

        /// <param name="columns">number of columns</param>

        /// <returns></returns>

        private static Row CreateContent(UInt32 index, DataRow dr, int columns)

        {

            Row objRow = new Row();

            Cell objCell;

            try

            {

                objRow.RowIndex = index;

                for (int i = 65; i < 65 + columns; i++)

                {

                    objCell = new Cell();

                    objCell.StyleIndex = 5;

                    objCell.DataType = CellValues.String;

                    objCell.CellReference = (char)i + index.ToString();

                    objCell.CellValue = new CellValue(dr.ItemArray[i - 65].ToString());

                    objRow.Append(objCell);

                }

            }

            catch (Exception ex)

            {

                throw;

            }

            return objRow;

        }

        /// <summary>

        /// Defines the Style sheet for the excel.

        /// </summary>

        /// <returns>Stylesheet</returns>

        private static Stylesheet CreateStylesheet()

        {

            Stylesheet ss = new Stylesheet();

            Fonts fts = new Fonts();

            DocumentFormat.OpenXml.Spreadsheet.Font ft = new DocumentFormat.OpenXml.Spreadsheet.Font();

            FontName ftn = new FontName();

            ftn.Val = StringValue.FromString("Calibri");

            DocumentFormat.OpenXml.Spreadsheet.FontSize ftsz = new DocumentFormat.OpenXml.Spreadsheet.FontSize();

            ftsz.Val = DoubleValue.FromDouble(11);

            ft.FontName = ftn;

            ft.FontSize = ftsz;

            fts.Append(ft);

            ft = new DocumentFormat.OpenXml.Spreadsheet.Font();

            ftn = new FontName();

            ftn.Val = StringValue.FromString("Palatino Linotype");

            ftsz = new DocumentFormat.OpenXml.Spreadsheet.FontSize();

            ftsz.Val = DoubleValue.FromDouble(18);

            ft.FontName = ftn;

            ft.FontSize = ftsz;

            fts.Append(ft);

            fts.Count = UInt32Value.FromUInt32((uint)fts.ChildElements.Count);

            Fills fills = new Fills();

            Fill fill;

            PatternFill patternFill;

            fill = new Fill();

            patternFill = new PatternFill();

            patternFill.PatternType = PatternValues.None;

            fill.PatternFill = patternFill;

            fills.Append(fill);

            fill = new Fill();

            patternFill = new PatternFill();

            patternFill.PatternType = PatternValues.Gray125;

            fill.PatternFill = patternFill;

            fills.Append(fill);

            fill = new Fill();

            patternFill = new PatternFill();

            patternFill.PatternType = PatternValues.Solid;

            patternFill.ForegroundColor = new ForegroundColor();

            patternFill.ForegroundColor.Rgb = HexBinaryValue.FromString("CDCDCD");

            patternFill.BackgroundColor = new BackgroundColor();

            patternFill.BackgroundColor.Rgb = patternFill.ForegroundColor.Rgb;

            fill.PatternFill = patternFill;

            fills.Append(fill);

            fills.Count = UInt32Value.FromUInt32((uint)fills.ChildElements.Count);

            Borders borders = new Borders();

            Border border = new Border();

            border.LeftBorder = new LeftBorder();

            border.RightBorder = new RightBorder();

            border.TopBorder = new TopBorder();

            border.BottomBorder = new BottomBorder();

            border.DiagonalBorder = new DiagonalBorder();

            borders.Append(border);

            border = new Border();

            border.LeftBorder = new LeftBorder();

            border.LeftBorder.Style = BorderStyleValues.Thin;

            border.RightBorder = new RightBorder();

            border.RightBorder.Style = BorderStyleValues.Thin;

            border.TopBorder = new TopBorder();

            border.TopBorder.Style = BorderStyleValues.Thin;

            border.BottomBorder = new BottomBorder();

            border.BottomBorder.Style = BorderStyleValues.Thin;

            border.DiagonalBorder = new DiagonalBorder();

            borders.Append(border);

            borders.Count = UInt32Value.FromUInt32((uint)borders.ChildElements.Count);

            CellStyleFormats csfs = new CellStyleFormats();

            CellFormat cf = new CellFormat();

            cf.NumberFormatId = 0;

            cf.FontId = 0;

            cf.FillId = 0;

            cf.BorderId = 0;

            csfs.Append(cf);

            csfs.Count = UInt32Value.FromUInt32((uint)csfs.ChildElements.Count);

            uint iExcelIndex = 164;

            NumberingFormats nfs = new NumberingFormats();

            CellFormats cfs = new CellFormats();

            NumberingFormat nfForcedText = new NumberingFormat();

            nfForcedText.NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++);

            nfForcedText.FormatCode = StringValue.FromString("@");

            nfs.Append(nfForcedText);

            cf = new CellFormat();

            cf.FontId = 0;

            cf.FillId = 0;

            cf.BorderId = 0;

            cf.FormatId = 0;

            cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);

            cfs.Append(cf);

            cf = new CellFormat();

            cf.FontId = 0;

            cf.FillId = 0;

            cf.BorderId = 1;

            cf.FormatId = 0;

            cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);

            cfs.Append(cf);

            cf = new CellFormat();

            cf.FontId = 0;

            cf.FillId = 0;

            cf.BorderId = 0;

            cf.FormatId = 0;

            cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);

            cfs.Append(cf);

            cf = new CellFormat();

            cf.NumberFormatId = nfForcedText.NumberFormatId;

            cf.FontId = 0;

            cf.FillId = 0;

            cf.BorderId = 0;

            cf.FormatId = 0;

            cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);

            cfs.Append(cf);

            cf = new CellFormat();

            cf.NumberFormatId = nfForcedText.NumberFormatId;

            cf.FontId = 1;

            cf.FillId = 0;

            cf.BorderId = 0;

            cf.FormatId = 0;

            cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);

            cfs.Append(cf);

            cf = new CellFormat();

            cf.FontId = 0;

            cf.FillId = 0;

            cf.BorderId = 1;

            cf.FormatId = 0;

            cfs.Append(cf);

            cf = new CellFormat();

            cf.FontId = 0;

            cf.FillId = 2;

            cf.BorderId = 1;

            cf.FormatId = 0;

            cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);

            cfs.Append(cf);

            cf = new CellFormat();

            cf.NumberFormatId = nfForcedText.NumberFormatId;

            cf.FontId = 0;

            cf.FillId = 2;

            cf.BorderId = 1;

            cf.FormatId = 0;

            cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);

            cfs.Append(cf);

            ss.Append(nfs);

            ss.Append(fts);

            ss.Append(fills);

            ss.Append(borders);

            ss.Append(csfs);

            ss.Append(cfs);

            CellStyles css = new CellStyles();

            CellStyle cs = new CellStyle();

            cs.Name = StringValue.FromString("Normal");

            cs.FormatId = 0;

            cs.BuiltinId = 0;

            css.Append(cs);

            css.Count = UInt32Value.FromUInt32((uint)css.ChildElements.Count);

            ss.Append(css);

            DifferentialFormats dfs = new DifferentialFormats();

            dfs.Count = 0;

            ss.Append(dfs);

            TableStyles tss = new TableStyles();

            tss.Count = 0;

            tss.DefaultTableStyle = StringValue.FromString("TableStyleMedium9");

            tss.DefaultPivotStyle = StringValue.FromString("PivotStyleLight16");

            ss.Append(tss);

            return ss;

        }

        /// <summary>

        /// Create the header row provided with the row index and header list with each column size

        /// </summary>

        /// <param name="index">row index</param>

        /// <param name="headerList">header name and column width</param>

        /// <returns></returns>

        private static Row CreateColumnHeader(UInt32 index, Dictionary<string, int> headerList)

        {

            Row objRow = new Row();

            objRow.RowIndex = index;

            Cell objCell;

            for (int i = 0; i < headerList.Count(); i++)

            {

                objCell = new Cell();

                objCell.DataType = CellValues.String;

                objCell.StyleIndex = 6;

                objCell.CellReference = Convert.ToChar(65 + i) + index.ToString();

                objCell.CellValue = new CellValue(headerList.Keys.ElementAt(i).ToString());

                objRow.Append(objCell);

            }

            return objRow;

        }

        /// <summary>

        /// Creating the filters based on which data row is generated. For creating header templates

        /// </summary>

        /// <param name="index">Row index</param>

        /// <param name="filters">Search key and its respective value</param>

        /// <returns></returns>

        private static Row CreateFilters(UInt32 index, Dictionary<string, string> filters)

        {

            Row objRow = new Row();

            try

            {

                objRow.RowIndex = index;

                Cell objcell;

                objcell = new Cell();

                objcell.DataType = CellValues.String;

                objcell.StyleIndex = 6;

                objcell.CellReference = "A" +index.ToString();

                objcell.CellValue = new CellValue(filters.Keys.ElementAt(Convert.ToInt32(index - 2)).ToString());

                objRow.Append(objcell);

                objcell = new Cell();

                objcell.DataType = CellValues.String;

                objcell.StyleIndex = 5;

                objcell.CellReference = "B" +index.ToString();

                objcell.CellValue = new CellValue(Convert.ToString(filters.ElementAt(Convert.ToInt32(index - 2)).Value));

                objRow.Append(objcell);

            }

            catch (Exception ex)

            {

                throw;

            }

            return objRow;

        }

    }
}
