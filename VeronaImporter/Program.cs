using System;
using System.Data;
using Microsoft.Data.SqlClient;
using OfficeOpenXml;
using System.IO;

class Program
{
    static void Main(string[] args)
    {
        string excelFilePath = @"C:\Users\Administrator\Desktop\Power\Verona\3 - GER 11 2024_finalTESTE.xlsx";
        string connectionString = "Server=hoh2k2203.hmv.org.br;Database=OnBaseDev;User Id=hsi;Password=wstinol;TrustServerCertificate=True;Connection Timeout=900;";
        string tableName = "VeronaTESTE";

        try
        {
            DataTable dataTable = ReadExcelFile(excelFilePath);

            // Create the table if it doesn't exist
            CreateTableIfNotExists(dataTable, connectionString, tableName);

            // Bulk insert data into the table
            BulkInsertToDatabase(dataTable, connectionString, tableName);
            Console.WriteLine("Data import completed successfully!");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    static DataTable ReadExcelFile(string filePath)
    {
        var dataTable = new DataTable();
        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            var worksheet = package.Workbook.Worksheets[0]; // Get the first worksheet
            int rows = worksheet.Dimension.Rows;
            int columns = worksheet.Dimension.Columns;

            // Add columns to the DataTable
            for (int col = 1; col <= columns; col++)
            {
                string columnName = worksheet.Cells[1, col].Text.Trim();
                if (!string.IsNullOrEmpty(columnName))
                {
                    dataTable.Columns.Add(columnName, typeof(string)); // Default to string
                }
            }

            // Add rows to the DataTable
            for (int row = 2; row <= rows; row++)
            {
                var dataRow = dataTable.NewRow();
                for (int col = 1; col <= columns; col++)
                {
                    dataRow[col - 1] = worksheet.Cells[row, col].Text;
                }
                dataTable.Rows.Add(dataRow);
            }
        }

        return dataTable;
    }

    static void CreateTableIfNotExists(DataTable dataTable, string connectionString, string tableName)
    {
        using (SqlConnection connection = new SqlConnection(connectionString))
        {
            connection.Open();

            // Check if the table exists
            string checkTableQuery = $@"
                IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = '{tableName}')
                BEGIN
                    CREATE TABLE {tableName} ({GenerateCreateTableColumns(dataTable)});
                END";

            using (SqlCommand command = new SqlCommand(checkTableQuery, connection))
            {
                command.ExecuteNonQuery();
            }
        }
    }

    static string GenerateCreateTableColumns(DataTable dataTable)
    {
        var columnDefinitions = new System.Text.StringBuilder();

        foreach (DataColumn column in dataTable.Columns)
        {
            columnDefinitions.Append($"[{column.ColumnName}] NVARCHAR(MAX), ");
        }

        // Remove the trailing comma and space
        columnDefinitions.Length -= 2;

        return columnDefinitions.ToString();
    }

    static void BulkInsertToDatabase(DataTable dataTable, string connectionString, string tableName)
    {
        using (SqlConnection connection = new SqlConnection(connectionString))
        {
            connection.Open();
            using (SqlBulkCopy bulkCopy = new SqlBulkCopy(connection))
            {
                bulkCopy.DestinationTableName = tableName;

                bulkCopy.BulkCopyTimeout = 900;

                // Map columns dynamically
                foreach (DataColumn column in dataTable.Columns)
                {
                    bulkCopy.ColumnMappings.Add(column.ColumnName, column.ColumnName);
                }

                bulkCopy.WriteToServer(dataTable);
            }
        }
    }
}
