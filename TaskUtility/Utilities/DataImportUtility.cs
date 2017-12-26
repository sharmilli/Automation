using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Configuration;

namespace TaskUtility.Utilities
{
    public static class DataImportUtility
    {
        /// <summary>
        /// Import all excel sheet data to a datatable
        /// </summary>
        /// <param name="folderPath"></param>
        /// <param name="fileName"></param>
        /// <param name="sheetName"></param>
        /// <param name="startColumn"></param>
        /// <param name="endColumn"></param>
        /// <param name="startRow"></param>
        /// <param name="endRow"></param>
        /// <param name="TableName"></param>
        /// <returns></returns>
        public static DataTable ExcelToDatatable(string folderPath, string fileName, string sheetName, string startColumn, string endColumn, string startRow, string endRow, string TableName)
        {
            DataTable dt = new DataTable();

            string fileFullPath = getFulFilePath(folderPath, fileName);
            string ConStr;
            ConStr = Constants.ExcelConnection;
            fileFullPath = fileFullPath + ".xlsx";
            ConStr = ConStr.Replace("<%dynamicPathComesHere%>", fileFullPath.ToString());
            var sheetname = sheetName + "$";
            OleDbConnection cnn = new OleDbConnection(ConStr);
            cnn.Open();

            DataTable dtOledb = new DataTable();
            String SQLDataQuery = Constants.SheetDataSQLQuery;

            if (startColumn == "" && endColumn == "" && startRow == "" && endRow == "")
            {
                SQLDataQuery = SQLDataQuery.Replace("<%SheetNameComesHere%>", sheetname);
            }
            else if ((startColumn != "" && endColumn != "") && (startRow == "" && endRow == ""))
            {
                var endcol = new String(endColumn.Where(Char.IsLetter).ToArray());
                SQLDataQuery = SQLDataQuery.Replace("<%SheetNameComesHere%>", sheetname + startColumn + ":" + endcol + "100000");
            }
            else
            {
                var strcol = new String(startColumn.Where(Char.IsLetter).ToArray());
                var endcol = new String(endColumn.Where(Char.IsLetter).ToArray());
                var strrow = Regex.Match(startRow, @"\d+").Value;
                var endrow = Regex.Match(endRow, @"\d+").Value;
                SQLDataQuery = SQLDataQuery.Replace("<%SheetNameComesHere%>", sheetname + startColumn + ":" + endcol + endrow);
            }

            OleDbCommand cmd = new OleDbCommand(SQLDataQuery, cnn);
            OleDbDataAdapter adp = new OleDbDataAdapter(cmd);
            adp.Fill(dtOledb);
            return dtOledb;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="FolderPath"></param>
        /// <param name="FileName"></param>
        /// <returns></returns>
        private static string getFulFilePath(string FolderPath, string FileName)
        {
            StringBuilder fullPath = new StringBuilder();
            fullPath.Append(FolderPath);
            fullPath.Append("\\");
            fullPath.Append(DateTime.Now.Year);
            fullPath.Append("\\");
            fullPath.Append(Enum.GetName(typeof(SpanishMonths), DateTime.Now.Month));
            fullPath.Append("\\");
            fullPath.Append(FileName);
            fullPath.Append(DateTime.Now.Year);
            fullPath.Append(Enum.GetName(typeof(SpanishMonths), DateTime.Now.Month));

            return fullPath.ToString();
        }

        /// <summary>
        /// Build Create table query
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="sheetname"></param>
        /// <param name="FileName"></param>
        /// <param name="TableName"></param>
        /// <param name="SchemaName"></param>
        /// <returns></returns>
        private static string GetCreateTableQuery(DataTable dt, String sheetname, string FileName, string TableName, string SchemaName)
        {
            sheetname = sheetname.Replace("$", "");
            StringBuilder tableDDL = new StringBuilder("");
            tableDDL.Append(Constants.ifExistsQuery);
            tableDDL.Append(Constants.checkObjectQuery1);
            tableDDL.Append(TableName);
            tableDDL.Append(Constants.checkObjectQuery2);
            tableDDL.Append(Constants.dropTableQuery);
            tableDDL.Append(TableName);
            tableDDL.Append(Constants.closingSquareBrace);
            tableDDL.Append(Constants.createTableQuery);
            tableDDL.Append(TableName);
            tableDDL.Append(Constants.closingSquareBrace);
            tableDDL.Append(Constants.openingRoundBrace);
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                if (i != dt.Columns.Count - 1)
                {
                    tableDDL.Append(Constants.openingSquareBrace);
                    tableDDL.Append(dt.Columns[i].ColumnName);
                    tableDDL.Append(Constants.closingSquareBrace + " ");
                    tableDDL.Append(Constants.Varchar);
                    tableDDL.Append(Constants.comma);
                }

                else
                {
                    tableDDL.Append(Constants.openingSquareBrace);
                    tableDDL.Append(dt.Columns[i].ColumnName);
                    tableDDL.Append(Constants.closingSquareBrace + " ");
                    tableDDL.Append(Constants.Varchar);
                }

            }
            tableDDL.Append(Constants.closingRoundBrace);
            tableDDL.Replace("<%SchemaNameComesHere%>", SchemaName);
            return tableDDL.ToString();

        }
        
        /// <summary>
        /// Create Table in DB
        /// </summary>
        /// <param name="SQLquery"></param>
        public static void CreateTable(string SQLquery)
        {
            string connectionString = GetConnectionString();
            DataTable dt = new DataTable();
            using (var connection = new SqlConnection(connectionString))
            {
                string queryString = SQLquery;
                var command = new SqlCommand(queryString, connection);
                connection.Open();
                command.ExecuteNonQuery();

            }
        }

        /// <summary>
        /// Sql Bulk Copy to import excel data to created table
        /// </summary>
        /// <param name="SQLConnection"></param>
        /// <param name="filename"></param>
        /// <param name="sheetname"></param>
        /// <param name="dt"></param>
        /// <param name="TableName"></param>
        /// <param name="SchemaName"></param>
        public static void SQLBulkCopy(string SQLConnection, string filename, string sheetname, DataTable dt, string TableName, string SchemaName)
        {
            SqlBulkCopy blk = new SqlBulkCopy(SQLConnection);

            StringBuilder destTableName = new StringBuilder("[" + SchemaName + "].");
            destTableName.Append(Constants.openingSquareBrace);
            destTableName.Append(TableName);
            destTableName.Append(Constants.closingSquareBrace);
            blk.DestinationTableName = destTableName.ToString();
            blk.WriteToServer(dt);

        }

        
        /// <summary>
        /// Getting connection details from configuration file
        /// </summary>
        /// <returns></returns>
        private static string GetConnectionString()
        {
            var connection = ConfigurationManager.ConnectionStrings["SQLConnection"].ConnectionString;
            return connection;
        }

        /// <summary>
        /// Fetch the data table from excel
        /// </summary>
        /// <returns></returns>
        public static DataTable GetDataTableFromExcel()
        {
            DataTable dt = new DataTable();
            var connectionString = GetConnectionString();
            using (var connection = new SqlConnection(connectionString))
            {
                string queryString = Constants.QueryToGetExceldetails;
                var command = new SqlCommand(queryString, connection);
                connection.Open();
                SqlDataAdapter da = new SqlDataAdapter(command);
                da.Fill(dt);
            }
            return dt;
        }
    }
}
