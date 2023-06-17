using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using System.Text;
using System.Xml.Linq;

public static class Excel2Json
{
    private class ExcelRule
    {
        public string excelName;
        public string[] sheetNames;
        public string className;
    }

    //======水平======
    const int LEFT = 1;
    const int ID = 2;

    //======垂直======
    const int TOP = 1;
    const int CHS = 2;
    const int NAME = 3;
    const int TYPE = 4;

    private static XElement root;

    private static string ns;
    private static string excel;
    private static string[] jsons;
    private static string[] csharps;


    private static List<ExcelRule> rules = new List<ExcelRule>();
    private static Dictionary<string, List<string>> classes = new Dictionary<string, List<string>>();
    private static Dictionary<string, JObject> tables = new Dictionary<string, JObject>();

    private static List<string> accesses = new List<string>();

    private static string manager = "TableManager";

    public static void Exprot(string export)
    {
        root = XDocument.Load(export).Root;

        GetPath();
        GetExcels();
        GetNameSpace();

        Clear();

        CollectExcels();
        ExportExcels();
        ExportManager();

        ShowAllAccess();
    }

    private static void ShowAllAccess()
    {
        Console.WriteLine("==============================================");
        foreach (var access in accesses)
        {
            ShowAccess(access);
        }
        Console.WriteLine("==============================================");
    }

    private static void ShowAccess(string access)
    {
        Console.WriteLine($"========{access}");
    }

    private static void Clear()
    {
        foreach (var csharp in csharps)
        {
            if (!Directory.Exists(csharp)) continue;
            DirectoryInfo info = new DirectoryInfo(csharp);
            info.Delete(true);
        }
        foreach (var json in jsons)
        {
            if (!Directory.Exists(json)) continue;
            DirectoryInfo info = new DirectoryInfo(json);
            info.Delete(true);
        }

    }

    private static void GetPath()
    {
        var path = root.Element("path");
        excel = path.Element("excel")?.Value;
        jsons = path.Elements("json")?.Select(xml => xml?.Value).ToArray();
        csharps = path.Elements("csharp")?.Select(xml => xml?.Value).ToArray();
    }

    private static void GetExcels()
    {
        var xExcels = root.Elements("excel");
        foreach (var excel in xExcels)
        {
            var xRules = excel.Elements("rule");
            var excelName = excel.Element("name");
            foreach (var rule in xRules)
            {
                var sheetNames = rule.Element("sheets");
                var className = rule.Element("class");
                rules.Add(new ExcelRule()
                {
                    excelName = excelName?.Value,
                    sheetNames = sheetNames?.Value.Split(","),
                    className = className?.Value,
                });
            }
        }
    }

    private static void GetNameSpace()
    {
        ns = root.Element("namespace")?.Value;
    }

    private static void CollectExcels()
    {
        foreach (var rule in rules)
        {
            CollectExcel(rule);
        }
    }

    private static void ExportExcels()
    {
        ExportClasses();
        ExportTables();
    }

    private static void ExportClasses()
    {
        foreach (var pair in classes)
        {
            if (CheckContents(pair.Value))
            {
                accesses.Add(pair.Key);
                foreach (var csharp in csharps)
                {
                    if (!Directory.Exists(csharp))
                    {
                        Directory.CreateDirectory(csharp);
                    }
                    File.WriteAllText($"{csharp}/{pair.Key}.cs", pair.Value[0]);
                }
            }
        }
    }

    private static void ExportTables()
    {
        foreach (var table in tables)
        {
            if (!accesses.Contains(table.Key))
            {
                Console.WriteLine($"==============================================");
                Console.WriteLine($"{table.Key}字段不完全匹配!!!!!!!->拒绝输出!!!!!!!!");
                continue;
            }
            foreach (var json in jsons)
            {
                if (!Directory.Exists(json))
                {
                    Directory.CreateDirectory(json);
                }
                File.WriteAllText($"{json}/{table.Key}.json", table.Value.ToString());
            }
        }
    }

    private static bool CheckContents(List<string> contents)
    {
        if (contents.Count <= 1) return true;
        foreach (var content in contents)
        {
            if (content != contents[0]) return false;
        }
        return true;
    }

    private static void CollectExcel(ExcelRule rule)
    {
        try
        {
            var excelPath = $"{excel}/{rule.excelName}";
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var stream = new FileStream(excelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (ExcelPackage excelPackage = new ExcelPackage(stream))
                {
                    Console.WriteLine($"==============================================");
                    Console.WriteLine($"开始收集->{rule.className}---{rule.excelName}");

                    JObject table;
                    ExcelWorksheet sheet;

                    foreach (var sheetName in rule.sheetNames)
                    {
                        sheet = GetSheet(excelPackage, sheetName);
                        if (!tables.ContainsKey(rule.className))
                        {
                            tables.Add(rule.className, new JObject());
                        }
                        table = tables[rule.className];
                        CollectClass(sheet, rule.className);
                        CollectTable(sheet, table);
                        Console.WriteLine($"正在收集->{rule.className}---{sheetName}");
                    }
                    Console.WriteLine($"收集完成->{rule.className}---{rule.excelName}");
                }
            }
        }
        catch (IOException ioe) { }
    }

    private static void CollectClass(ExcelWorksheet sheet, string className)
    {
        var id = sheet.Cells[TYPE, ID].GetValue<string>();
        var name = className;
        var col = sheet.Dimension.End.Column;
        List<List<string>> builds = new List<List<string>>();

        for (int j = 2; j <= col; j++)
        {
            var build = new List<string>();
            // 为*不收集次字段
            if (sheet.Cells[TOP, j].GetValue<string>() == "*") continue;

            if (sheet.Cells[CHS, j] == null) continue;

            // 字段收集
            build.Add(sheet.Cells[CHS, j].GetValue<string>());
            build.Add(sheet.Cells[NAME, j].GetValue<string>());
            build.Add(sheet.Cells[TYPE, j].GetValue<string>());

            builds.Add(build);
        }

        StringBuilder builder = new StringBuilder();
        var level = 0;

        builder.AppendLine("using Newtonsoft.Json;");
        builder.AppendLine("using System.Collections.Generic;");
        builder.AppendLine(null);

        if (!string.IsNullOrEmpty(ns))
        {
            builder.AppendLine($"namespace {ns}");
            builder.AppendLine($"{Indent(level)}{{");
        }
        builder.AppendLine($"{Indent(++level)}/// <summary>");
        builder.AppendLine($"{Indent(level)}/// {sheet.Cells[1, 1].GetValue<string>()}");
        builder.AppendLine($"{Indent(level)}/// <summary>");
        builder.AppendLine($"{Indent(level)}public class {name}");
        builder.AppendLine($"{Indent(level)}{{");
        level++;
        foreach (var build in builds)
        {
            builder.AppendLine($"{Indent(level)}/// <summary>");
            builder.AppendLine($"{Indent(level)}/// {build[0]}");
            builder.AppendLine($"{Indent(level)}/// </summary>");
            builder.AppendLine($"{Indent(level)}public {TrueType(build[2])} {build[1]} {{ get; set; }}");
            builder.AppendLine(null);
        }

        builder.AppendLine($"{Indent(level)}private static Dictionary<{id}, {name}> table;");
        builder.AppendLine(null);
        builder.AppendLine($"{Indent(level)}public static {name} Get({id} id)");
        builder.AppendLine($"{Indent(level)}{{");
        builder.AppendLine($"{Indent(++level)}table.TryGetValue(id, out {name} value);");
        builder.AppendLine($"{Indent(level)}return value;");
        builder.AppendLine($"{Indent(--level)}}}");

        builder.AppendLine(null);
        builder.AppendLine($"{Indent(level)}static {name}()");
        builder.AppendLine($"{Indent(level)}{{");
        builder.AppendLine($"{Indent(++level)}table = JsonConvert.DeserializeObject<Dictionary<{id}, {name}>>({manager}.Load(typeof({name})));");
        builder.AppendLine($"{Indent(--level)}}}");


        builder.AppendLine($"{Indent(--level)}}}");
        if (!string.IsNullOrEmpty(ns))
        {
            builder.AppendLine($"{Indent(--level)}}}");
        }

        if (!classes.ContainsKey(className))
        {
            classes.Add(className, new List<string>());
        }
        classes[className].Add(builder.ToString());
    }

    public static ExcelWorksheet GetSheet(ExcelPackage excelPackage, string sheetName)
    {
        foreach (var worksheet in excelPackage.Workbook.Worksheets)
        {
            if (worksheet.Name == sheetName) return worksheet;
        }
        return null;
    }

    private static void ExportManager()
    {
        StringBuilder builder = new StringBuilder();
        var level = 0;
        builder.AppendLine("using System;");
        builder.AppendLine(null);

        if (!string.IsNullOrEmpty(ns))
        {
            builder.AppendLine($"namespace {ns}");
            builder.AppendLine($"{Indent(level)}{{");
        }

        builder.AppendLine($"{Indent(++level)}public static class {manager}");
        builder.AppendLine($"{Indent(level)}{{");
        builder.AppendLine($"{Indent(++level)}public static Func<string, string> onLoad {{ get; set; }}");
        builder.AppendLine($"{Indent(level)}public static string Load(Type type)");
        builder.AppendLine($"{Indent(level)}{{");
        builder.AppendLine($"{Indent(++level)}return onLoad?.Invoke(type.Name);");
        builder.AppendLine($"{Indent(--level)}}}");
        builder.AppendLine($"{Indent(--level)}}}");

        if (!string.IsNullOrEmpty(ns))
        {
            builder.AppendLine($"{Indent(--level)}}}");
        }

        foreach (var csharp in csharps)
        {
            if (!Directory.Exists(csharp))
            {
                Directory.CreateDirectory(csharp);
            }
            File.WriteAllText($"{csharp}/{manager}.cs", builder.ToString());
        }
    }

    private static void CollectTable(ExcelWorksheet sheet, JObject table)
    {
        var row = sheet.Dimension.End.Row;
        var col = sheet.Dimension.End.Column;

        // 收集每一行的数据
        for (int i = TYPE + 1; i <= row; i++)
        {
            if (sheet.Cells[i, LEFT].GetValue<string>() == "*") continue;
            JObject jobject = new JObject();
            var name = sheet.Cells[i, ID].GetValue<string>();
            if (name == null) continue;
            for (int j = 2; j <= col; j++)
            {
                if (sheet.Cells[TOP, j].GetValue<string>() == "*") continue;
                if (sheet.Cells[CHS, j] == null) continue;
                var type = sheet.Cells[TYPE, j].GetValue<string>();

                switch (type)
                {
                    case "int":
                        {
                            AddObject(jobject, sheet.Cells[NAME, j].GetValue<string>(), sheet.Cells[i, j].GetValue<string>(), (val) => string.IsNullOrEmpty(val) ? 0 : int.Parse(val));
                            break;
                        }
                    case "float":
                        {
                            AddObject(jobject, sheet.Cells[NAME, j].GetValue<string>(), sheet.Cells[i, j].GetValue<string>(), (val) => string.IsNullOrEmpty(val) ? 0 : float.Parse(val));
                            break;
                        }
                    case "string":
                        {
                            AddObject(jobject, sheet.Cells[NAME, j].GetValue<string>(), sheet.Cells[i, j].GetValue<string>(), (val) => (val));
                            break;
                        }
                    case "int[]":
                        {
                            AddArray(jobject, sheet.Cells[NAME, j].GetValue<string>(), sheet.Cells[i, j].GetValue<string>(), (val) => int.Parse(val));
                            break;
                        }
                    case "float[]":
                        {
                            AddArray(jobject, sheet.Cells[NAME, j].GetValue<string>(), sheet.Cells[i, j].GetValue<string>(), (val) => float.Parse(val));
                            break;
                        }
                    case "string[]":
                        {
                            AddArray(jobject, sheet.Cells[NAME, j].GetValue<string>(), sheet.Cells[i, j].GetValue<string>(), (val) => val);
                            break;
                        }
                    case "int[][]":
                        {
                            AddArrayArray(jobject, sheet.Cells[NAME, j].GetValue<string>(), sheet.Cells[i, j].GetValue<string>(), (val) => int.Parse(val));
                            break;
                        }
                    case "float[][]":
                        {
                            AddArrayArray(jobject, sheet.Cells[NAME, j].GetValue<string>(), sheet.Cells[i, j].GetValue<string>(), (val) => float.Parse(val));
                            break;
                        }
                    case "string[][]":
                        {
                            AddArrayArray(jobject, sheet.Cells[NAME, j].GetValue<string>(), sheet.Cells[i, j].GetValue<string>(), (val) => val);
                            break;
                        }
                    case "int:int":
                        {
                            AddDictionary(jobject, sheet.Cells[NAME, j].GetValue<string>(), sheet.Cells[i, j].GetValue<string>(), (val) => int.Parse(val));
                            break;
                        }
                    case "int:string":
                        {
                            AddDictionary(jobject, sheet.Cells[NAME, j].GetValue<string>(), sheet.Cells[i, j].GetValue<string>(), (val) => val);
                            break;
                        }
                    case "string:int":
                        {
                            AddDictionary(jobject, sheet.Cells[NAME, j].GetValue<string>(), sheet.Cells[i, j].GetValue<string>(), (val) => int.Parse(val));
                            break;
                        }
                    case "string:string":
                        {
                            AddDictionary(jobject, sheet.Cells[NAME, j].GetValue<string>(), sheet.Cells[i, j].GetValue<string>(), (val) => val);
                            break;
                        }
                    case "int:int[]":
                        {
                            AddArrayDictionary(jobject, sheet.Cells[NAME, j].GetValue<string>(), sheet.Cells[i, j].GetValue<string>(), (val) => int.Parse(val));
                            break;
                        }
                    case "int:float[]":
                        {
                            AddArrayDictionary(jobject, sheet.Cells[NAME, j].GetValue<string>(), sheet.Cells[i, j].GetValue<string>(), (val) => float.Parse(val));
                            break;
                        }
                    case "int:string[]":
                        {
                            AddArrayDictionary(jobject, sheet.Cells[NAME, j].GetValue<string>(), sheet.Cells[i, j].GetValue<string>(), (val) => val);
                            break;
                        }
                    case "string:int[]":
                        {
                            AddArrayDictionary(jobject, sheet.Cells[NAME, j].GetValue<string>(), sheet.Cells[i, j].GetValue<string>(), (val) => int.Parse(val));
                            break;
                        }
                    case "string:float[]":
                        {
                            AddArrayDictionary(jobject, sheet.Cells[NAME, j].GetValue<string>(), sheet.Cells[i, j].GetValue<string>(), (val) => float.Parse(val));
                            break;
                        }
                    case "string:string[]":
                        {
                            AddArrayDictionary(jobject, sheet.Cells[NAME, j].GetValue<string>(), sheet.Cells[i, j].GetValue<string>(), (val) => val);
                            break;
                        }
                }
            }
            table.Add(name, jobject);
        }
    }

    private static string Indent(int level)
    {
        var sb = new StringBuilder();
        for (int i = 0; i < level; i++)
        {
            sb.Append("\t");
        }
        return sb.ToString();
    }

    private static void AddObject(JObject jobject, string name, string value, Func<string, JToken> func)
    {
        value = (value == null) ? "" : value;
        jobject.Add(name, func(value));
    }

    private static void AddArray(JObject jobject, string name, string value, Func<string, JToken> func)
    {
        JArray jarray0 = new JArray();
        string[] elements = (value == null) ? new string[0] : value.Split(',');
        foreach (var element in elements)
        {
            jarray0.Add(func(element));
        }
        jobject.Add(name, jarray0);
    }

    private static void AddArrayArray(JObject jobject, string name, string value, Func<string, JToken> func)
    {
        JArray jarray0 = new JArray();
        string[] elements = (value == null) ? new string[0] : value.Substring(0, value.Length - 1).Split("],");
        foreach (var element in elements)
        {
            JArray array = new JArray();
            string arrayStr = element.Substring(1);
            var datas = arrayStr.Split(',');
            foreach (var data in datas)
            {
                array.Add(func(data));
            }
            jarray0.Add(array);
        }
        jobject.Add(name, jarray0);
    }

    private static void AddDictionary(JObject jobject, string name, string value, Func<string, JToken> func)
    {
        JObject table = new JObject();
        string[] elements = (value == null) ? new string[0] : value.Split(',');
        foreach (var element in elements)
        {
            string[] pair = element.Split(':');
            table.Add(pair[0], func(pair[1]));
        }
        jobject.Add(name, table);
    }

    private static void AddArrayDictionary(JObject jobject, string name, string value, Func<string, JToken> func)
    {
        JObject table = new JObject();
        string[] elements = (value == null) ? new string[0] : value.Substring(0, value.Length - 1).Split("],");
        foreach (var element in elements)
        {
            JArray array = new JArray();
            string[] pair = element.Split(':');
            string arrayStr = pair[1].Substring(1);
            var datas = arrayStr.Split(',');
            foreach (var data in datas)
            {
                array.Add(func(data));
            }
            table.Add(pair[0], array);
        }
        jobject.Add(name, table);
    }

    public static string TrueType(string type)
    {
        if (!type.Contains(':')) return type;
        var pair = type.Split(':');
        return $"Dictionary<{pair[0]},{pair[1]}>";
    }
}
