using CommandLine;
using ExdExport;
using Lumina;
using Lumina.Data;
using Lumina.Data.Structs.Excel;
using Lumina.Excel;
using Lumina.Excel.Sheets.Experimental;
using Lumina.Text.ReadOnly;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Unicode;

#pragma warning disable PendingExcelSchema

var options = Parser.Default.ParseArguments<Options>(args).Value;

if (string.IsNullOrWhiteSpace(options.Path))
    throw new Exception("Path to sqpack directory is empty.");

if (options.Languages == null || options.Languages.Length == 0)
    throw new Exception("No languages selected.");

var version = File.ReadAllText(Path.Join(options.Path, "../ffxivgame.ver")).TrimEnd();

Console.WriteLine($"Path: {options.Path}");
Console.WriteLine($"Version: {version}");
Console.WriteLine($"Selected languages: {string.Join(", ", options.Languages ?? [])}");

var gameData = new GameData(options.Path, new()
{
    PanicOnSheetChecksumMismatch = false,
    CacheFileResources = false,
});

var sheetTypes = new List<Type>(
    Assembly.GetAssembly(typeof(Achievement))!.GetTypes()
        .Where(type => type.Namespace == "Lumina.Excel.Sheets.Experimental" && !type.IsNested));

var typeNameCache = new Dictionary<Type, string>();

foreach (var selectedLangStr in options.Languages!)
{
    foreach (var (lang, langStr) in LanguageUtil.LanguageMap)
    {
        if (langStr == selectedLangStr)
        {
            ProcessLanguage(lang);
            break;
        }
    }
}

Console.WriteLine("Done!");

// ---------------------------

void ProcessLanguage(Language language)
{
    var langStr = LanguageUtil.GetLanguageStr(language);

    foreach (var sheetName in gameData.Excel.SheetNames)
    {
        var sheetOutPath = Path.Join(options.ExportPath, $"/{version}/{langStr}/{sheetName}.json");
        var dirPath = sheetOutPath[0..sheetOutPath.LastIndexOf('/')];
        if (!Directory.Exists(dirPath))
            Directory.CreateDirectory(dirPath);

        if (File.Exists(sheetOutPath) && new FileInfo(sheetOutPath).Length > 0)
            continue;

        var rowType = sheetTypes.Find(type => type.Name == sheetName);
        if (rowType != null)
        {
            if (rowType.GetInterfaces()[0].GetGenericTypeDefinition().IsAssignableTo(typeof(IExcelSubrow<>)))
            {
                var sheet = gameData.Excel.GetType().GetMethod("GetSubrowSheet")?.MakeGenericMethod(rowType)?.Invoke(gameData.Excel, [language, sheetName]);
                if (sheet == null)
                    continue;

                Console.WriteLine($"[{langStr}] Processing {sheetName}");
                using var file = File.OpenWrite(sheetOutPath);
                ProcessGeneratedSubrowSheet(sheet, sheetName, rowType, file);
                continue;
            }
            else
            {
                var sheet = gameData.Excel.GetType().GetMethod("GetSheet")?.MakeGenericMethod(rowType)?.Invoke(gameData.Excel, [language, sheetName]);
                if (sheet == null)
                    continue;

                Console.WriteLine($"[{langStr}] Processing {sheetName}");
                using var file = File.OpenWrite(sheetOutPath);
                ProcessGeneratedSheet(sheet, sheetName, rowType, file);
                continue;
            }
        }

        try
        {
            var sheet = gameData.Excel.GetSheet<RawRow>(language, sheetName);
            if (sheet == null)
                continue;

            Console.WriteLine($"[{langStr}] Processing {sheetName}");
            using var file = File.OpenWrite(sheetOutPath);
            ProcessSheet(sheet, sheetName, rowType, file);
            continue;
        }
        catch (Exception)
        {
            try
            {
                var sheet = gameData.Excel.GetSubrowSheet<RawSubrow>(language, sheetName);
                if (sheet == null)
                    continue;

                Console.WriteLine($"[{langStr}] Processing {sheetName}");
                using var file = File.OpenWrite(sheetOutPath);
                ProcessSubrowSheet(sheet, sheetName, rowType, file);
                continue;
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}

void ProcessGeneratedSheet(object sheet, string sheetName, Type rowType, FileStream fileStream)
{
    using var writer = new Utf8JsonWriter(fileStream, new()
    {
        Indented = true,
        Encoder = JavaScriptEncoder.Create(UnicodeRanges.BasicLatin, UnicodeRanges.Latin1Supplement)
    });

    var sheetType = sheet.GetType();
    var columns = (IReadOnlyList<ExcelColumnDefinition>)sheetType.GetProperty("Columns")!.GetValue(sheet)!;
    var rowCount = (int)sheetType.GetProperty("Count")!.GetValue(sheet)!;

    writer.WriteStartObject();
    writer.WritePropertyName("meta");
    writer.WriteStartObject();
    writer.WriteString("sheetName", sheetName);
    writer.WriteNumber("numColumns", columns.Count);
    writer.WriteNumber("numRows", rowCount);

    var i = 0;
    using var pb = new ProgressBar();

    var rawSeStringData = rowType.Name is "CustomTalkDefineClient" or "QuestDefineClient";

    writer.WriteEndObject(); // meta
    writer.WritePropertyName("rows");
    writer.WriteStartArray();

    var props = rowType.GetProperties(BindingFlags.DeclaredOnly | BindingFlags.Instance | BindingFlags.Public);

    foreach (var row in (dynamic)sheet)
    {
        pb.Report((double)i / rowCount);
        i++;

        writer.WriteStartObject();
        writer.WriteNumber("@rowId", row.RowId);

        foreach (var prop in props)
        {
            if (prop.Name == "RowId")
                continue;

            writer.WritePropertyName(prop.Name);
            WriteValue(writer, prop.PropertyType, prop.GetValue(row), rawSeStringData);
            // writer.WriteCommentValue(); for index and offset?
        }

        writer.WriteEndObject();
    }

    writer.WriteEndArray(); // rows
    writer.WriteEndObject();
    writer.Flush();
}

void ProcessGeneratedSubrowSheet(object sheet, string sheetName, Type rowType, FileStream fileStream)
{
    using var writer = new Utf8JsonWriter(fileStream, new()
    {
        Indented = true,
        Encoder = JavaScriptEncoder.Create(UnicodeRanges.BasicLatin, UnicodeRanges.Latin1Supplement)
    });

    var sheetType = sheet.GetType();
    var columns = (IReadOnlyList<ExcelColumnDefinition>)sheetType.GetProperty("Columns")!.GetValue(sheet)!;
    var rowCount = (int)sheetType.GetProperty("Count")!.GetValue(sheet)!;

    writer.WriteStartObject();
    writer.WritePropertyName("meta");
    writer.WriteStartObject();
    writer.WriteString("sheetName", sheetName);
    writer.WriteNumber("numColumns", columns.Count);
    writer.WriteNumber("numRows", rowCount);

    var i = 0;
    using var pb = new ProgressBar();

    var rawSeStringData = rowType.Name is "CustomTalkDefineClient" or "QuestDefineClient";

    writer.WriteEndObject(); // meta
    writer.WritePropertyName("rows");
    writer.WriteStartArray();

    var props = rowType.GetProperties(BindingFlags.DeclaredOnly | BindingFlags.Instance | BindingFlags.Public);

    foreach (var row in (dynamic)sheet)
    {
        pb.Report((double)i / rowCount);
        i++;

        writer.WriteStartObject();
        writer.WriteNumber("@rowId", row.RowId);
        writer.WritePropertyName("subrows");
        writer.WriteStartArray();

        foreach (var subrow in row)
        {
            pb.Report((double)i / rowCount);
            i++;

            writer.WriteStartObject();
            writer.WriteNumber("@subRowId", subrow.SubrowId);

            foreach (var prop in props)
            {
                if (prop.Name is "RowId" or "SubrowId")
                    continue;

                writer.WritePropertyName(prop.Name);
                WriteValue(writer, prop.PropertyType, prop.GetValue(subrow), rawSeStringData);
                // writer.WriteCommentValue(); for index and offset?
            }

            writer.WriteEndObject();
        }

        writer.WriteEndArray(); // subrows
        writer.WriteEndObject();
    }

    writer.WriteEndArray(); // rows
    writer.WriteEndObject();
    writer.Flush();
}

void ProcessSheet(ExcelSheet<RawRow> sheet, string sheetName, Type? sheetType, FileStream fileStream)
{
    using var writer = new Utf8JsonWriter(fileStream, new()
    {
        Indented = true,
        Encoder = JavaScriptEncoder.Create(UnicodeRanges.BasicLatin, UnicodeRanges.Latin1Supplement)
    });

    writer.WriteStartObject();
    writer.WritePropertyName("meta");
    writer.WriteStartObject();
    writer.WriteString("sheetName", sheetName);
    writer.WriteNumber("numColumns", sheet.Columns.Count);
    writer.WriteNumber("numRows", sheet.Count);

    var i = 0;
    using var pb = new ProgressBar();

    if (!sheet.GetType().IsGenericType)
    {
        if (sheetType == null)
            goto WriteRawSheet;

        var rawSeStringData = sheetType.Name is "CustomTalkDefineClient" or "QuestDefineClient";

        object objExcelSheet = typeof(GameData)
            .GetMethods(BindingFlags.Public | BindingFlags.Instance)
            .Where(m => m.Name == "GetExcelSheet")
            .First(m => m.GetParameters().Length == 1)
            .MakeGenericMethod(sheetType)
            .Invoke(gameData, [sheet.Language])!;

        if (objExcelSheet == null)
            goto WriteRawSheet;

        writer.WriteEndObject(); // meta
        writer.WritePropertyName("rows");
        writer.WriteStartArray();

        dynamic sheetInstance = Convert.ChangeType(objExcelSheet, typeof(ExcelSheet<>).MakeGenericType(sheetType));
        PropertyInfo[]? props = null;

        foreach (var row in sheetInstance)
        {
            pb.Report((double)i / sheetInstance.RowCount);
            i++;

            props ??= row.GetType().GetProperties(BindingFlags.DeclaredOnly | BindingFlags.Instance | BindingFlags.Public);

            writer.WriteStartObject();
            writer.WriteNumber("@rowId", row.RowId);
            writer.WriteNumber("@subRowId", row.SubRowId);

            foreach (var prop in props)
            {
                writer.WritePropertyName(prop.Name);
                WriteValue(writer, prop.PropertyType, prop.GetValue(row), rawSeStringData);
                // writer.WriteCommentValue(); for index and offset?
            }

            writer.WriteEndObject();
        }

        writer.WriteEndArray(); // rows
        writer.WriteEndObject();
        writer.Flush();
        return;
    }

WriteRawSheet:
    writer.WritePropertyName("columns");
    writer.WriteStartArray();
    foreach (var column in sheet.Columns)
        writer.WriteStringValue(column.Type.ToString());
    writer.WriteEndArray();

    writer.WriteEndObject(); // meta
    writer.WritePropertyName("rows");
    writer.WriteStartArray();
    foreach (var row in sheet)
    {
        pb.Report((double)i / sheet.Count);
        i++;

        writer.WriteStartObject();
        writer.WriteNumber("@rowId", row.RowId);
        // writer.WriteNumber("@subRowId", row.SubRowId);
        writer.WritePropertyName("columns");
        writer.WriteStartArray();

        for (var j = 0; j < sheet.Columns.Count; j++)
        {
            var value = row.ReadColumn(j);
            WriteValue(writer, value!.GetType(), value);
        }

        writer.WriteEndArray(); // columns
        writer.WriteEndObject(); // row
    }

    writer.WriteEndArray(); // rows
    writer.WriteEndObject();
    writer.Flush();
    return;
}

void ProcessSubrowSheet(SubrowExcelSheet<RawSubrow> sheet, string sheetName, Type? sheetType, FileStream fileStream)
{
    using var writer = new Utf8JsonWriter(fileStream, new()
    {
        Indented = true,
        Encoder = JavaScriptEncoder.Create(UnicodeRanges.BasicLatin, UnicodeRanges.Latin1Supplement)
    });

    writer.WriteStartObject();
    writer.WritePropertyName("meta");
    writer.WriteStartObject();
    writer.WriteString("sheetName", sheetName);
    writer.WriteNumber("numColumns", sheet.Columns.Count);
    writer.WriteNumber("numRows", sheet.Count);

    var i = 0;
    using var pb = new ProgressBar();

    if (!sheet.GetType().IsGenericType)
    {
        if (sheetType == null)
            goto WriteRawSheet;

        var rawSeStringData = sheetType.Name is "CustomTalkDefineClient" or "QuestDefineClient";

        object objExcelSheet = typeof(GameData)
            .GetMethods(BindingFlags.Public | BindingFlags.Instance)
            .Where(m => m.Name == "GetExcelSheet")
            .First(m => m.GetParameters().Length == 1)
            .MakeGenericMethod(sheetType)
            .Invoke(gameData, [sheet.Language])!;

        if (objExcelSheet == null)
            goto WriteRawSheet;

        writer.WriteEndObject(); // meta
        writer.WritePropertyName("rows");
        writer.WriteStartArray();

        dynamic sheetInstance = Convert.ChangeType(objExcelSheet, typeof(ExcelSheet<>).MakeGenericType(sheetType));
        PropertyInfo[]? props = null;

        foreach (var row in sheetInstance)
        {
            pb.Report((double)i / sheetInstance.RowCount);
            i++;

            props ??= row.GetType().GetProperties(BindingFlags.DeclaredOnly | BindingFlags.Instance | BindingFlags.Public);

            writer.WriteStartObject();
            writer.WriteNumber("@rowId", row.RowId);
            writer.WriteNumber("@subRowId", row.SubRowId);

            foreach (var prop in props)
            {
                writer.WritePropertyName(prop.Name);
                WriteValue(writer, prop.PropertyType, prop.GetValue(row), rawSeStringData);
                // writer.WriteCommentValue(); for index and offset?
            }

            writer.WriteEndObject();
        }

        writer.WriteEndArray(); // rows
        writer.WriteEndObject();
        writer.Flush();
        return;
    }

WriteRawSheet:
    writer.WritePropertyName("columns");
    writer.WriteStartArray();
    foreach (var column in sheet.Columns)
        writer.WriteStringValue(column.Type.ToString());
    writer.WriteEndArray();

    writer.WriteEndObject(); // meta
    writer.WritePropertyName("rows");
    writer.WriteStartArray();
    foreach (var row in sheet)
    {
        pb.Report((double)i / sheet.Count);
        i++;

        writer.WriteStartObject();
        writer.WriteNumber("@rowId", row.RowId);

        writer.WritePropertyName("subrows");
        writer.WriteStartArray();

        foreach (var subrow in row)
        {
            writer.WriteStartObject();
            writer.WriteNumber("@subRowId", subrow.SubrowId);
            writer.WritePropertyName("columns");
            writer.WriteStartArray();

            for (var j = 0; j < sheet.Columns.Count; j++)
            {
                var value = subrow.ReadColumn(j);
                WriteValue(writer, value!.GetType(), value);
            }

            writer.WriteEndArray(); // columns
            writer.WriteEndObject(); // subrow
        }

        writer.WriteEndArray(); // subrows
        writer.WriteEndObject(); // row
    }

    writer.WriteEndArray(); // rows
    writer.WriteEndObject();
    writer.Flush();
    return;
}

void WriteValue(Utf8JsonWriter writer, Type type, object? value, bool rawSeStringData = false)
{
    if (value == null)
    {
        writer.WriteNullValue();
    }
    else if (value is ReadOnlySeString seString)
    {
        writer.WriteStringValue(seString.ToMacroString());
    }
    else if (type.IsPrimitive)
    {
        switch (value)
        {
            case bool val:
                writer.WriteBooleanValue(val);
                break;
            case byte val:
                writer.WriteNumberValue(val);
                break;
            case sbyte val:
                writer.WriteNumberValue(val);
                break;
            case short val:
                writer.WriteNumberValue(val);
                break;
            case ushort val:
                writer.WriteNumberValue(val);
                break;
            case int val:
                writer.WriteNumberValue(val);
                break;
            case uint val:
                writer.WriteNumberValue(val);
                break;
            case long val:
                writer.WriteNumberValue(val);
                break;
            case ulong val:
                writer.WriteNumberValue(val);
                break;
            case char val:
                writer.WriteStringValue(val.ToString());
                break;
            case double val:
                writer.WriteNumberValue(val);
                break;
            case float val:
                writer.WriteNumberValue(val);
                break;
            case string val:
                writer.WriteStringValue(val);
                break;
            default:
                throw new Exception($"Unhandled primitive type: {type.Name}");
        }
    }
    else if (type.IsGenericType && type.GetGenericTypeDefinition().IsAssignableTo(typeof(Collection<>)))
    {
        writer.WriteStartArray();

        var elementType = type.GenericTypeArguments[0];
        foreach (var item in (IEnumerable)value)
        {
            WriteValue(writer, elementType!, item);
        }

        writer.WriteEndArray();
    }
    else if (type.IsGenericType && type.GetGenericTypeDefinition().IsAssignableTo(typeof(SubrowCollection<>)))
    {
        writer.WriteStartArray();

        var elementType = type.GenericTypeArguments[0];
        foreach (var item in (IEnumerable)value)
        {
            WriteValue(writer, elementType!, item);
        }

        writer.WriteEndArray();
    }
    else if (value is RowRef rowRef)
    {
        writer.WriteStartObject();
        writer.WriteNumber("@rowId", rowRef.RowId);
        writer.WriteEndObject();
        return;
    }
    else if (type.IsGenericType && (type.GetGenericTypeDefinition().IsAssignableTo(typeof(RowRef<>)) || type.GetGenericTypeDefinition().IsAssignableTo(typeof(SubrowRef<>))))
    {
        var valueRowType = type.GenericTypeArguments[0];
        var valueRowId = type.GetProperty("RowId")!.GetValue(value)!;

        writer.WriteStartObject();
        writer.WriteString("@type", valueRowType.Name ?? "");
        writer.WriteNumber("@rowId", (uint)valueRowId);
        writer.WriteEndObject();
        return;
    }
    else if (type.IsValueType) // structs
    {
        writer.WriteStartObject();

        var props = type.GetProperties(BindingFlags.DeclaredOnly | BindingFlags.Instance | BindingFlags.Public);

        foreach (var prop in props)
        {
            writer.WritePropertyName(prop.Name);
            WriteValue(writer, prop.PropertyType, prop.GetValue(value, null));
        }

        writer.WriteEndObject();
    }
    else
    {
        throw new Exception($"Unhandled type {type.FullName}");
    }
}

public class Options
{
    [Option("path", HelpText = "The path to the sqpack directory.", Default = @"C:\Program Files (x86)\SquareEnix\FINAL FANTASY XIV - A Realm Reborn\game\sqpack")]
    public string? Path { get; set; }

    [Option("outdir", HelpText = "The path to the output directory.", Default = "export")]
    public string? ExportPath { get; set; }

    [Option('l', "languages", HelpText = "The languages to export.", Default = new string[] { "de", "en", "fr", "ja" }, Separator = ',')]
    public string[]? Languages { get; set; }
}
