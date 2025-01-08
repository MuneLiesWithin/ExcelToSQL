using System;
using System.Data;
using Microsoft.Data.SqlClient;
using OfficeOpenXml;
using System.IO;
using Microsoft.Extensions.Configuration;

class Program
{
    static void Main(string[] args)
    {
        Console.WriteLine("Starting the import process..."+DateTime.Now);

        // Load configuration
        Console.WriteLine("Loading configuration...");
        IConfiguration configuration = new ConfigurationBuilder()
            .SetBasePath(AppContext.BaseDirectory)
            .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
            .Build();

        // Get settings from configuration
        string? excelFilePath = configuration["FilePath"];
        string? connectionString = configuration.GetConnectionString("DefaultConnection");
        string? tableName = configuration["TableName"];

        if (string.IsNullOrWhiteSpace(excelFilePath))
        {
            Console.WriteLine("Error: File path is not provided in the configuration.");
            return;
        }
        Console.WriteLine($"Excel file path: {excelFilePath}");

        if (string.IsNullOrWhiteSpace(connectionString))
        {
            Console.WriteLine("Error: Connection string is not provided in the configuration.");
            return;
        }
        Console.WriteLine("Connection string loaded successfully.");

        if (string.IsNullOrWhiteSpace(tableName))
        {
            Console.WriteLine("Error: Table name is not provided in the configuration.");
            return;
        }
        Console.WriteLine($"Target table: {tableName}");

        try
        {
            // Read data from the Excel file
            Console.WriteLine("Reading data from the Excel file...");
            DataTable dataTable = ReadExcelFile(excelFilePath);
            Console.WriteLine($"Excel file read successfully. Rows: {dataTable.Rows.Count}, Columns: {dataTable.Columns.Count}");

            // Create the table if it doesn't exist
            Console.WriteLine("Checking if the table exists and creating it if necessary...");
            CreateTableIfNotExists(dataTable, connectionString, tableName);
            Console.WriteLine("Table checked/created successfully.");

            // Bulk insert data into the table
            Console.WriteLine("Starting bulk insert into the database...");
            BulkInsertToDatabase(dataTable, connectionString, tableName);
            Console.WriteLine("Bulk insert completed successfully.");

            Console.WriteLine("Data import process completed successfully!"+DateTime.Now);
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

            Console.WriteLine($"Reading Excel file. Rows: {rows}, Columns: {columns}");

            // Add columns to the DataTable
            for (int col = 1; col <= columns; col++)
            {
                string columnName = worksheet.Cells[1, col].Text.Trim();
                if (!string.IsNullOrEmpty(columnName))
                {
                    dataTable.Columns.Add(columnName, typeof(string)); // Default to string
                    Console.WriteLine($"Added column: {columnName}");
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
            Console.WriteLine($"Checking if table '{tableName}' exists...");

            // Check if the table exists
            string checkTableQuery = $@"
                IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = '{tableName}')
                BEGIN
                    CREATE TABLE {tableName} ({GenerateCreateTableColumns(dataTable)});
                END";

            using (SqlCommand command = new SqlCommand(checkTableQuery, connection))
            {
                command.ExecuteNonQuery();
                Console.WriteLine($"Table '{tableName}' checked/created successfully.");
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
            Console.WriteLine($"Starting bulk insert into '{tableName}'...");

            using (SqlBulkCopy bulkCopy = new SqlBulkCopy(connection))
            {
                bulkCopy.DestinationTableName = tableName;
                bulkCopy.BulkCopyTimeout = 900;

                // Map columns dynamically
                foreach (DataColumn column in dataTable.Columns)
                {
                    bulkCopy.ColumnMappings.Add(column.ColumnName, column.ColumnName);
                    Console.WriteLine($"Mapped column: {column.ColumnName}");
                }

                bulkCopy.WriteToServer(dataTable);
                Console.WriteLine("Bulk insert operation completed.");
            }
        }
    }
}
