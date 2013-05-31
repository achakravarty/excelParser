namespace ExcelParser.Core
{
    internal static class ConnectionStringBuilder
    {
        private const string OleDbConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0 Xml;HDR=YES\"";

        internal static string BuildConnectionString(string fileName, ExcelConnectionType connectionType)
        {
            switch (connectionType)
            {
                case ExcelConnectionType.OleDb: return string.Format(OleDbConnectionString, fileName);
                default: return string.Format(OleDbConnectionString, fileName);
            }
        }
    }
}