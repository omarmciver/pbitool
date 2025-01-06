using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using Microsoft.AnalysisServices.AdomdClient;

namespace PBIXAnalyzer
{

    public class PBIXAnalyzer
    {
    private readonly string _pbixPath;

    public PBIXAnalyzer(string pbixPath)
    {
        if (!File.Exists(pbixPath))
            throw new FileNotFoundException($"PBIX file not found: {pbixPath}");

        if (!Path.GetExtension(pbixPath).Equals(".pbix", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException($"File must have .pbix extension, got: {Path.GetExtension(pbixPath)}");

        _pbixPath = pbixPath;
    }

    public void PrintContents()
    {
        Console.WriteLine($"\nContents of {Path.GetFileName(_pbixPath)}:");
        Console.WriteLine(new string('-', 80));

        var contents = ListContents();
        foreach (var item in contents)
        {
            Console.WriteLine($"{item.Name,-60} {item.Size,15:N0} bytes");
        }

        Console.WriteLine($"\nTotal files: {contents.Count}");
    }

    private List<(string Name, long Size)> ListContents()
    {
        var contents = new List<(string Name, long Size)>();
        using var archive = ZipFile.OpenRead(_pbixPath);
        
        foreach (var entry in archive.Entries)
        {
            contents.Add((entry.FullName, entry.Length));
        }

        return contents;
    }

    public void AnalyzeModel()
    {
        Console.WriteLine("\nAnalyzing Power BI Data Model...");
        
        var modelInfo = ExtractDataModel();
        if (modelInfo == null)
        {
            Console.WriteLine("Failed to analyze DataModel");
            return;
        }

        // Print model information
        PrintTables(modelInfo.Tables);
        PrintRelationships(modelInfo.Relationships);
        PrintMeasures(modelInfo.Measures);
    }

    private ModelInfo? ExtractDataModel()
    {
        var tempDir = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
        Directory.CreateDirectory(tempDir);

        try
        {
            using (var archive = ZipFile.OpenRead(_pbixPath))
            {
                var dataModelEntry = archive.GetEntry("DataModel");
                if (dataModelEntry != null)
                {
                    var dataModelPath = Path.Combine(tempDir, "DataModel");
                    dataModelEntry.ExtractToFile(dataModelPath);
                }
            }

            var connectionString = $"Provider=MSOLAP;Data Source=$Embedded$;Initial Catalog=Microsoft_SQLServer_AnalysisServices;Locale Identifier=1033;Persist Security Info=True;Location={tempDir}\\DataModel";
            
            var modelInfo = new ModelInfo();
            using (var connection = new AdomdConnection(connectionString))
            {
                connection.Open();
                ExtractTables(connection, modelInfo);
                ExtractRelationships(connection, modelInfo);
                ExtractMeasures(connection, modelInfo);
            }

            return modelInfo;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error analyzing DataModel: {ex.Message}");
            return null;
        }
        finally
        {
            Directory.Delete(tempDir, true);
        }
    }

    private void ExtractTables(AdomdConnection connection, ModelInfo modelInfo)
    {
        var tablesCommand = connection.CreateCommand();
        tablesCommand.CommandText = @"
            SELECT 
                [TABLE_NAME],
                [DESCRIPTION],
                [TABLE_TYPE]
            FROM $SYSTEM.DBSCHEMA_TABLES
            WHERE TABLE_TYPE = 'TABLE'";

        using var reader = tablesCommand.ExecuteReader();
        while (reader.Read())
        {
            var tableInfo = new TableInfo
            {
                Name = reader.GetString(0),
                Description = reader.IsDBNull(1) ? null : reader.GetString(1),
                Type = reader.GetString(2)
            };

            // Get columns for the table
            var columnsCommand = connection.CreateCommand();
            columnsCommand.CommandText = $@"
                SELECT 
                    [COLUMN_NAME],
                    [DATA_TYPE],
                    [DESCRIPTION]
                FROM $SYSTEM.DBSCHEMA_COLUMNS 
                WHERE TABLE_NAME = '{tableInfo.Name}'";

            using var columnsReader = columnsCommand.ExecuteReader();
            while (columnsReader.Read())
            {
                tableInfo.Columns.Add(new ColumnInfo
                {
                    Name = columnsReader.GetString(0),
                    DataType = columnsReader.GetString(1),
                    Description = columnsReader.IsDBNull(2) ? null : columnsReader.GetString(2)
                });
            }

            modelInfo.Tables.Add(tableInfo);
        }
    }

    private void ExtractRelationships(AdomdConnection connection, ModelInfo modelInfo)
    {
        var command = connection.CreateCommand();
        command.CommandText = @"
            SELECT 
                [PK_TABLE_NAME],
                [PK_COLUMN_NAME],
                [FK_TABLE_NAME],
                [FK_COLUMN_NAME]
            FROM $SYSTEM.DBSCHEMA_RELATIONSHIPS";

        using var reader = command.ExecuteReader();
        while (reader.Read())
        {
            modelInfo.Relationships.Add(new RelationshipInfo
            {
                PkTable = reader.GetString(0),
                PkColumn = reader.GetString(1),
                FkTable = reader.GetString(2),
                FkColumn = reader.GetString(3)
            });
        }
    }

    private void ExtractMeasures(AdomdConnection connection, ModelInfo modelInfo)
    {
        var command = connection.CreateCommand();
        command.CommandText = @"
            SELECT 
                [MEASURE_NAME],
                [MEASURE_CAPTION],
                [EXPRESSION],
                [MEASURE_IS_VISIBLE]
            FROM $SYSTEM.MDSCHEMA_MEASURES
            WHERE MEASURE_IS_VISIBLE";

        using var reader = command.ExecuteReader();
        while (reader.Read())
        {
            modelInfo.Measures.Add(new MeasureInfo
            {
                Name = reader.GetString(0),
                Caption = reader.GetString(1),
                Expression = reader.GetString(2),
                IsVisible = reader.GetBoolean(3)
            });
        }
    }

    private void PrintTables(List<TableInfo> tables)
    {
        Console.WriteLine("\nTables:");
        Console.WriteLine(new string('-', 80));
        
        foreach (var table in tables)
        {
            Console.WriteLine($"\nTable: {table.Name}");
            Console.WriteLine($"Type: {table.Type}");
            if (!string.IsNullOrEmpty(table.Description))
                Console.WriteLine($"Description: {table.Description}");

            Console.WriteLine("\nColumns:");
            foreach (var column in table.Columns)
            {
                Console.WriteLine($"- {column.Name} ({column.DataType})");
                if (!string.IsNullOrEmpty(column.Description))
                    Console.WriteLine($"  Description: {column.Description}");
            }
        }
    }

    private void PrintRelationships(List<RelationshipInfo> relationships)
    {
        Console.WriteLine("\nRelationships:");
        Console.WriteLine(new string('-', 80));
        
        foreach (var rel in relationships)
        {
            Console.WriteLine($"{rel.PkTable}.{rel.PkColumn} -> {rel.FkTable}.{rel.FkColumn}");
        }
    }

    private void PrintMeasures(List<MeasureInfo> measures)
    {
        Console.WriteLine("\nMeasures:");
        Console.WriteLine(new string('-', 80));
        
        foreach (var measure in measures)
        {
            Console.WriteLine($"\nMeasure: {measure.Name}");
            Console.WriteLine($"Caption: {measure.Caption}");
            Console.WriteLine($"Expression: {measure.Expression}");
        }
    }
}

public class ModelInfo
{
    public List<TableInfo> Tables { get; } = new();
    public List<RelationshipInfo> Relationships { get; } = new();
    public List<MeasureInfo> Measures { get; } = new();
}

public class TableInfo
{
    public string Name { get; set; } = "";
    public string? Description { get; set; }
    public string Type { get; set; } = "";
    public List<ColumnInfo> Columns { get; } = new();
}

public class ColumnInfo
{
    public string Name { get; set; } = "";
    public string DataType { get; set; } = "";
    public string? Description { get; set; }
}

public class RelationshipInfo
{
    public string PkTable { get; set; } = "";
    public string PkColumn { get; set; } = "";
    public string FkTable { get; set; } = "";
    public string FkColumn { get; set; } = "";
}

public class MeasureInfo
{
    public string Name { get; set; } = "";
    public string Caption { get; set; } = "";
    public string Expression { get; set; } = "";
    public bool IsVisible { get; set; }
}
