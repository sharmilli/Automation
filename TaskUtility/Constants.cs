namespace TaskUtility
{

    public enum Freequency
    {
        Daily = 1,
        Weekly = 2,
        Monthly = 3,
        Quaterly = 4

    }
    public enum SpanishMonths
    {
        enero = 1,
        febrero = 2,
        marzo = 3,
        abril = 4,
        mayo = 5,
        junio = 6,
        julio = 7,
        agosto = 8,
        septiembre = 9,
        octubre = 10,
        noviembre = 11,
        diciembre = 12
    }

    public enum TaskTypes
    {
        Email=1,
        ImportData=2,
        ExportData=3,
        MonitorFileshare=4,
        ExecuteProcedures=5
    }

    public static class Constants
    {
        public static string doubleSlash = "\\";

        public static string ExcelConnection = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= <%dynamicPathComesHere%> ;Extended Properties='Excel 8.0;HDR=YES'";

        public static string SheetDataSQLQuery = "select * from [<%SheetNameComesHere%>]";

        public static string ifExistsQuery = "IF EXISTS (SELECT * FROM sys.objects WHERE object_id = ";

        public static string checkObjectQuery1 = "OBJECT_ID(N'[<%SchemaNameComesHere%>].[";

        public static string underScore = "_";

        public static string checkObjectQuery2 = "]') AND type in (N'U'))";

        public static string dropTableQuery = "Drop Table [<%SchemaNameComesHere%>].[";

        public static string closingSquareBrace = "]";


        public static string createTableQuery = "Create table [<%SchemaNameComesHere%>].[";

        public static string openingRoundBrace = "(";

        public static string openingSquareBrace = "[";

        public static string Varchar = "NVarchar(max)";

        public static string comma = ",";

        public static string closingRoundBrace = ")";

        public static string QueryToGetExceldetails = "select * from ExcelDetails";//TODO: ?? hard coded
    }
}
