using CommandLine;
using ExdExport;
using Lumina;
using Lumina.Data;
using Lumina.Excel;
using Lumina.Excel.GeneratedSheets2;
using Lumina.Text;
using Lumina.Text.Payloads;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Unicode;

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
        .Where(type => !type.IsNested && (type.FullName!.StartsWith("Lumina.Excel.GeneratedSheets2") || type.FullName!.StartsWith("Lumina.Excel.CustomSheets"))));

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

    var removeSheetFromCacheMethodInfo = gameData.Excel.GetType()
        .GetMethods(BindingFlags.Public | BindingFlags.Instance)
        .First(methodInfo => methodInfo.Name == "RemoveSheetFromCache" && methodInfo.GetParameters().Length == 0);

    foreach (var sheetName in gameData.Excel.GetSheetNames())
    {
        if (sheetName == "None")
            continue;

        var sheetOutPath = Path.Join(options.ExportPath, $"/{version}/{langStr}/{sheetName}.json");
        var dirPath = sheetOutPath[0..sheetOutPath.LastIndexOf('/')];
        if (!Directory.Exists(dirPath))
            Directory.CreateDirectory(dirPath);

        if (File.Exists(sheetOutPath) && new FileInfo(sheetOutPath).Length > 0)
            continue;

        var sheet = gameData.Excel.GetSheetRaw(sheetName, language);
        if (sheet == null)
            continue;

        var sheetType = sheetTypes.Find(type => type.Name == sheet.Name);

        Console.WriteLine($"[{langStr}] Processing {sheetName}");
        using (var file = File.OpenWrite(sheetOutPath))
            ProcessSheet(sheet, sheetType, file);

        if (sheetType == null)
            continue;

        removeSheetFromCacheMethodInfo.MakeGenericMethod(sheetType).Invoke(gameData.Excel, null);
    }
}

void ProcessSheet(RawExcelSheet sheet, Type? sheetType, FileStream fileStream)
{
    using var writer = new Utf8JsonWriter(fileStream, new()
    {
        Indented = true,
        Encoder = JavaScriptEncoder.Create(UnicodeRanges.BasicLatin, UnicodeRanges.Latin1Supplement)
    });

    writer.WriteStartObject();
    writer.WritePropertyName("meta");
    writer.WriteStartObject();
    writer.WriteString("sheetName", sheet.Name);
    writer.WriteString("gameVersion", version);
    writer.WriteNumber("numColumns", sheet.ColumnCount);
    writer.WriteNumber("numRows", sheet.RowCount);

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
            .Invoke(gameData, [sheet.RequestedLanguage])!;

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
    foreach (var rowParser in sheet.GetRowParsers())
    {
        for (var j = 0; j < sheet.ColumnCount; j++)
        {
            writer.WriteStringValue(rowParser.ReadColumnRaw(j)!.GetType().Name);
        }

        break;
    }
    writer.WriteEndArray();

    writer.WriteEndObject(); // meta
    writer.WritePropertyName("rows");
    writer.WriteStartArray();
    foreach (var rowParser in sheet.GetRowParsers())
    {
        pb.Report((double)i / sheet.RowCount);
        i++;

        writer.WriteStartObject();
        writer.WriteNumber("@rowId", rowParser.RowId);
        writer.WriteNumber("@subRowId", rowParser.SubRowId);
        writer.WritePropertyName("columns");
        writer.WriteStartArray();

        for (var j = 0; j < sheet.ColumnCount; j++)
        {
            var value = rowParser.ReadColumnRaw(j);
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

void WriteValue(Utf8JsonWriter writer, Type type, object? value, bool rawSeStringData = false)
{
    if (value == null)
    {
        writer.WriteNullValue();
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
    else if (type.IsArray)
    {
        writer.WriteStartArray();

        var elementType = type.GetElementType();
        foreach (var item in (Array)value)
        {
            WriteValue(writer, elementType!, item);
        }

        writer.WriteEndArray();
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
    else if (value is ILazyRow lazyRow)
    {
        if (lazyRow.RawRow != null)
        {
            writer.WriteStartObject();
            writer.WriteString("@type", lazyRow.RawRow?.SheetName ?? "");
            writer.WriteNumber("@rowId", (int)lazyRow.Row);
            writer.WriteEndObject();
            return;
        }

        if (type.IsGenericType)
        {
            if (!typeNameCache.TryGetValue(type, out var typeName))
            {
                if (type.GetGenericArguments().Length > 0)
                {
                    var rowType = type.GetGenericArguments()[0];
                    if (rowType != typeof(ExcelRow))
                    {
                        typeName = rowType.Name;
                        typeNameCache.Add(type, typeName);
                    }
                }
            }

            if (!string.IsNullOrEmpty(typeName))
            {
                writer.WriteStartObject();
                writer.WriteString("@type", typeName ?? "");
                writer.WriteNumber("@rowId", (int)lazyRow.Row);
                writer.WriteEndObject();
                return;
            }
        }

        writer.WriteNumberValue((int)lazyRow.Row);
    }
    else if (value is SeString seString)
    {
        if (rawSeStringData)
        {
            writer.WriteStringValue(seString.RawData);
        }
        else
        {
            var sb = new StringBuilder();

            foreach (var payload in seString.Payloads)
            {
                switch (payload.PayloadType)
                {
                    case PayloadType.SoftHyphen:
                        continue;

                    case PayloadType.NewLine:
                        sb.Append('\n');
                        continue;

                    default:
                        sb.Append(payload.ToString());
                        break;
                }
            }

            writer.WriteStringValue(sb.ToString());
        }
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
