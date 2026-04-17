//#define DEBUG

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;

//行从0开始索引，列从0开始索引
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExportExcelConfig
{
    enum WorkType
    {
        None,
        ExportConfigXml,
        ExportConfigScript,
        ExportLocalization,
    }
    class Program
    {
        private const string ENUM_PREFIX = "Em";
        const string logPrefix = "【导出excel】";
        const string logIndent = "====";
        private static string baseClass = " : ConfBase";
        
        //配置
        private static string confScriptOutputDir = "";
        private static string confXmlOutputDir = "";
        private static string confExcelInputDir = "";
        private static string l10nSourceFile = "";
        private static string l10nOutputDir = "";
        private static string l10nOutputPrefix = "";//前缀+地区名.txt

        private static string help = @"
Params:
WorkType=ExportConfigXml/ExportConfigScript/ExportLocalization

ConfigExcelInputDir=xxx (for ExportConfigXml and ExportConfigScript)
ConfigScriptOutputDir=xxx (for ExportConfigScript)
ConfigXmlOutputDir=xxx (for ExportConfigXml)

LocalizationExcelFilePath=xxx (for ExportLocalization)
LocalizationOutputDir=xxx (for ExportLocalization)
LocalizationFileOutputPrefix=xxx (for ExportLocalization)
";

        private static WorkType workType;

        private static List<string> baseClassFields = new List<string>()
        {
            "id", "name",
        };

        private static void Main(string[] args)
        {
            //无参数显示帮助
            if(args.Length <= 1)
            {
#if DEBUG
                args = new[]
                {
                    "WorkType=ExportLocalization",

                    "ConfigExcelInputDir=../../TestDatas/",
                    "ConfigXmlOutputDir=../../TestDatas/Output/Xml/",
                    "ConfigScriptOutputDir=../../TestDatas/Output/Config/",

                    "LocalizationExcelFilePath=../../TestDatas/Localization.xlsx",
                    "LocalizationOutputDir=../../TestDatas/Output/Localization/",
                    "LocalizationFileOutputPrefix=L10n_",
                };
#else
                Console.WriteLine(help);
                PressAnyKeyToExist();
                return;
#endif
            }

            workType = WorkType.None;
            for(int i = 0; i < args.Length; i++)
            {
                var parts = args[i].Split(new char[]{'='});
                if (parts.Length != 2)
                {
                    Console.WriteLine($"error param format error：{args[i]}，should be：KEY=VALUE");
                    continue;
                }
                var key = parts[0];
                var value = parts[1];
                switch (key)
                {
                    case "WorkType":
                        switch (value)
                        {
                            case "ExportConfigXml":
                                workType = WorkType.ExportConfigXml;
                                break;
                            case "ExportConfigScript":
                                workType = WorkType.ExportConfigScript;
                                break;
                            case "ExportLocalization":
                                workType = WorkType.ExportLocalization;
                                break;
                            default:
                                Console.WriteLine("workType is wrong");
                                break;
                        }
                        break;

                    case "ConfigExcelInputDir":
                        confExcelInputDir = value;
                        break;
                    case "ConfigScriptOutputDir":
                        confScriptOutputDir = value;
                        break;
                    case "ConfigXmlOutputDir":
                        confXmlOutputDir = value;
                        break;
                    case "LocalizationExcelFilePath":
                        l10nSourceFile = value;
                        break;
                    case "LocalizationOutputDir":
                        l10nOutputDir = value;
                        break;
                    case "LocalizationFileOutputPrefix":
                        l10nOutputPrefix = value;
                        break;
                    default:
                        Console.WriteLine($"error: Unknown key [{key}]");
                        break;
                }
            }

            LogSettings();

            //创建目录
            try
            {
                switch(workType)
                {
                    case WorkType.ExportConfigScript:
                        Directory.CreateDirectory(confScriptOutputDir);
                        break;
                    case WorkType.ExportConfigXml:
                        Directory.CreateDirectory(confXmlOutputDir);
                        break;
                    case WorkType.ExportLocalization:
                        Directory.CreateDirectory(l10nOutputDir);
                        break;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                PressAnyKeyToExist();
                return;
            }

            //开始工作
            switch (workType)
            {
                case WorkType.ExportConfigXml:
                case WorkType.ExportConfigScript:
                    Work_XmlOrCs(workType);
                    break;
                case WorkType.ExportLocalization:
                    Work_Localization(workType);
                    break;
            }
            
            PressAnyKeyToExist();
        }

        private static void Work_XmlOrCs(WorkType workType)
        {
            var filePaths = Directory.GetFiles(confExcelInputDir, "*.xlsx");
            foreach (var path in filePaths)
            {
                IWorkbook workbook = null;
                try
                {
                    using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        var excelFileName = Path.GetFileName(path);
                        if(path.EndsWith(".xlsx")) // 2007
                        {
                            workbook = new XSSFWorkbook(fs);
                        }
                        else if (path.EndsWith(".xls")) // 2003
                        {
                            workbook = new HSSFWorkbook(fs);
                        }
                        if (workbook == null)
                        {
                            Console.WriteLine($"{logPrefix}{logIndent}读取文件失败: {excelFileName}");
                            continue;
                        }

                        //索引表
                        ISheet sheet = workbook.GetSheetAt(0);
                        //从第2行开始，读取要导出的表。一行一个。
                        for (int i = 1; i <= sheet.LastRowNum; i++)
                        {
                            var row = sheet.GetRow(i);
                            if(row == null)
                            {
                                Console.WriteLine($"Warning: {path} sheet0 row={i+1} has no data ");
                                continue;
                            }

                            var index = -1;
                            try
                            {
                                index = (int)row.GetCell(0).NumericCellValue;
                            }
                            catch (Exception e)
                            {
                                Console.WriteLine(e);
                            }
                            if (index <= 0)
                            {
                                Console.WriteLine($"{logPrefix}{logIndent}读取索引失败: {excelFileName}");
                                continue;
                            }

                            var fileName = row.GetCell(1).StringCellValue;
                            if (workType == WorkType.ExportConfigXml)
                            {
                                ExportXml(workbook.GetSheetAt(index), fileName, excelFileName);
                            }
                            else
                            {
                                ExportConfigFile(workbook.GetSheetAt(index), fileName, excelFileName);
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.ToString());
                }
            }

            Console.WriteLine();
            Console.WriteLine($"处理了{filePaths.Length}个文件，按任意键退出");
        }

        private static void Work_Localization(WorkType workType)
        {
            var n = 0;
            try
            {
                using (var fs = new FileStream(l10nSourceFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    IWorkbook workbook = null;
                    if (l10nSourceFile.EndsWith(".xlsx")) // 2007
                    {
                        workbook = new XSSFWorkbook(fs);
                    }
                    else if (l10nSourceFile.EndsWith(".xls")) // 2003
                    {
                        workbook = new HSSFWorkbook(fs);
                    }

                    //第一列是key，后面的列是地区名
                    ISheet sheet = workbook.GetSheetAt(0);
                    var row0 = sheet.GetRow(0);
                    var lastCellNum = row0.LastCellNum;
                    var lastRowNum = sheet.LastRowNum;
                    for (int i = 1; i < lastCellNum; i++)
                    {
                        string region = row0.Cells[i].StringCellValue;
                        string outputFileName = $"{l10nOutputPrefix}{region}.txt";
                        var sw = new StreamWriter(l10nOutputDir + outputFileName);
                        //一行一个，从第三行开始
                        for (int j = 2; j <= lastRowNum; j++)
                        {
                            var row = sheet.GetRow(j);
                            if (row == null)
                            {
                                Console.WriteLine($"Warning: row={j+1} has no data ");
                                continue;
                            }

                            try
                            {
                                sw.Write($"{row.Cells[0].StringCellValue} = {row.GetCell(i)?.StringCellValue}\n");
                            }
                            catch (Exception e)
                            {
                                Console.WriteLine($"{i+1}列 {j+1}行 <{region}> {row.GetCell(i)?.ToString() ?? "CellHasNoValue"}");
                                Console.WriteLine(e);
                            }
                        }
                        sw.Close();
                        n++;
                        Console.WriteLine($"导出: {outputFileName}");
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }

            Console.WriteLine();
            Console.WriteLine($"导出了{n}个文件，按任意键退出");
        }

        private static void LogSettings()
        {
            Console.WriteLine("---------------------------------------------------------------");
            if (workType == WorkType.ExportConfigXml || workType == WorkType.ExportConfigScript)
            {
                Console.WriteLine("Excel目录: " + confExcelInputDir);
                if (workType == WorkType.ExportConfigXml)
                {
                    Console.WriteLine("Xml输出目录: " + confXmlOutputDir);
                }
                else
                {
                    Console.WriteLine("Config输出目录: " + confScriptOutputDir);
                }
            }
            else if(workType == WorkType.ExportLocalization)
            {
                Console.WriteLine("本地化源文件: " + l10nSourceFile);
                Console.WriteLine("本地化输出目录: " + l10nOutputDir);
                Console.WriteLine("本地化输出前缀: " + l10nOutputPrefix);
            }
            Console.WriteLine("---------------------------------------------------------------");
        }

        //如果不存在，则创建
        private static void ExportConfigFile(ISheet sheet, string sheetName, string excelFileName)
        {
            string scriptName = GetConfScriptName(sheetName);
            string fileFullPath = $"{confScriptOutputDir}{scriptName}.cs";
            FileStream fs = new FileStream(fileFullPath, FileMode.Create, FileAccess.Write);
            StreamWriter sw = new StreamWriter(fs);
            Console.WriteLine($"{logPrefix}{logIndent}生成表脚本：{excelFileName}/{scriptName}");
            //类名
            sw.Write("//此文件由工具导出，手写会被覆盖。\n");
            sw.Write("using UnityEngine;\n\n[System.Serializable]\n");
            sw.Write($"public class {scriptName}{baseClass}\n{{\n");

            //字段
            var row0 = sheet.GetRow(0);
            var row1 = sheet.GetRow(1);
            var row2 = sheet.GetRow(2);
            var row3 = sheet.GetRow(3);
            var lastCellNum = row0.LastCellNum;
            for (int i = 0; i < lastCellNum; i++)
            {
                try
                {
                    //是否导出
                    if (row0.GetCell(i).ToString() != "1") continue;
                    //基类的字段名不写入
                    var fieldName = row2.GetCell(i).ToString();
                    if (baseClassFields.Contains(fieldName))
                    {
                        continue;
                    }

                    //注释
                    sw.Write($"\t//{row3.GetCell(i)}\n");
                    //声明
                    sw.Write($"\tpublic {row1.GetCell(i)} {fieldName};\n");
                }
                catch (Exception e)
                {
                    Console.WriteLine($"第{i}列的表头读取异常");
                    Console.WriteLine(e);
                }
            }
            sw.Write("}");
            sw.Close();
            fs.Close();
        }

        private static void ExportXml(ISheet sheet, string sheetName, string excelFileName)
        {
            string scriptName = GetConfScriptName(sheetName);
            string fileFullPath = $"{confXmlOutputDir}{sheetName}.xml";
            XmlTextWriter writer = new XmlTextWriter(fileFullPath, Encoding.UTF8);
            Console.WriteLine($"{logPrefix}{logIndent}导出XMl: {excelFileName}/{sheetName}");

            writer.Formatting = Formatting.Indented;
            writer.WriteStartDocument();
            writer.WriteStartElement("ArrayOf" + scriptName);

            var row0 = sheet.GetRow(0);
            var row1 = sheet.GetRow(1);
            var row2 = sheet.GetRow(2);
            var lastRowRum = sheet.LastRowNum;
            var lastCellNum = row0.LastCellNum;
            for (int i = 4; i <= lastRowRum; i++)
            {
                writer.WriteStartElement(scriptName);
                var curRow = sheet.GetRow(i);
                for (int j = 0; j <= lastCellNum; j++)
                {
                    var fieldName = row2.GetCell(j)?.ToString();
                    var fieldType = row1.GetCell(j)?.ToString();
                    var cell = curRow.GetCell(j);
                    ExportXmlField(writer, i, fieldName, fieldType, cell);
                }
                writer.WriteEndElement();
            }
            writer.WriteEndElement();
            writer.Close();
        }

        private static void ExportXmlField(XmlTextWriter writer, int row, string fieldName, string fieldType, ICell cell)
        {
            if (cell == null || fieldName == null || fieldType == null)
            {
                return;
            }

            writer.WriteStartElement(fieldName);
            //按使用频率依次处理
            if(fieldType == "int")
            {
                writer.WriteValue((int)cell.NumericCellValue);
            }
            else if(fieldType == "string")
            {
                writer.WriteValue(cell.StringCellValue);
            }
            else if(fieldType == "float")
            {
                writer.WriteValue(cell.NumericCellValue);
            }
            else if(fieldType == "Vector2")
            {
                var str = cell.StringCellValue;
                if(!string.IsNullOrEmpty(str))
                {
                    if (TryParseVector2(str, out Vector2 v))
                    {
                        v.WriteXml(writer);
                    }
                    else
                    {
                        LogGetValueFromString(row+1, fieldName, fieldType, str);
                    }
                }
            }
            else if(fieldType == "Vector3")
            {
                var str = cell.StringCellValue;
                if(!string.IsNullOrEmpty(str))
                {
                    if (TryParseVector3(str, out Vector3 v))
                    {
                        v.WriteXml(writer);
                    }
                    else
                    {
                        LogGetValueFromString(row+1, fieldName, fieldType, str);
                    }
                }
            }
            else if(fieldType == "bool")
            {
                var str = cell.StringCellValue;
                if(!string.IsNullOrEmpty(str))
                {
                    if (bool.TryParse(str, out bool b))
                    {
                        writer.WriteValue(b);
                    }
                    else
                    {
                        LogGetValueFromString(row+1, fieldName, fieldType, str);
                    }
                }
            }
            else if(fieldType == "long")
            {
                var str = cell.StringCellValue;
                if(!string.IsNullOrEmpty(str))
                {
                    if (long.TryParse(str, out long l))
                    {
                        writer.WriteValue(l);
                    }
                    else
                    {
                        LogGetValueFromString(row+1, fieldName, fieldType, str);
                    }
                }
            }
            else if(fieldType == "int[]")
            {
                var str = cell.StringCellValue;
                if(!string.IsNullOrEmpty(str))
                {
                    if (TryParseIntArray(str, out int[] array))
                    {
                        foreach (var i in array)
                        {
                            writer.WriteStartElement("int");
                            writer.WriteValue(i);
                            writer.WriteEndElement();
                        }
                    }
                    else
                    {
                        LogGetValueFromString(row+1, fieldName, fieldType, str);
                    }
                }
            }
            else if(fieldType.StartsWith(ENUM_PREFIX))
            {
                var str = cell.StringCellValue;
                if(!string.IsNullOrEmpty(str))
                {
                   writer.WriteValue(str);
                }
            }
            else
            {
                Console.WriteLine($"{logPrefix}从行{row + 1}的{fieldName}字段获取{fieldType}类型的值失败：不支持的数据类型 {cell}");
            }

            writer.WriteEndElement();
        }
        private static void LogGetValueFromString(int row, string valueName, string valueType, string valueString)
        {
            Console.WriteLine($"{logPrefix}从行{row + 1}的{valueName}字段获取{valueType}类型的值失败：数据格式错误 {valueString}");
        }

        private static bool TryParseVector2(string value, out Vector2 v)
        {
            v.x = 0;
            v.y = 0;
            int index = 0;
            int d = 1;
            for(; d<3; d++)
            {
                string floatString = GetFloatInString(value, ref index);
                if(!float.TryParse(floatString, out float f))
                    return false;
                if(d == 1) v.x = f;
                if(d == 2) v.y = f;
            }
            return d == 3;
        }

        private static bool TryParseVector3(string value, out Vector3 v)
        {
            v.x = 0;
            v.y = 0;
            v.z = 0;
            int index = 0;
            int d = 1;
            for(; d<4; d++)
            {
                string floatString = GetFloatInString(value, ref index);
                if(!float.TryParse(floatString, out float f))
                    return false;
                if(d == 1) v.x = f;
                if(d == 2) v.y = f;
                if(d == 3) v.z = f;
            }
            return d == 4;
        }

        private static string GetFloatInString(string value, ref int start)
        {
            int valueStart = -1;
            for(; start<value.Length; start++)
            {
                //先找到值的开始
                char c = value[start];
                if(valueStart < 0)
                {
                    if(('0'<=c && c<='9') || c == '-')
                        valueStart = start;
                }
                else
                {
                    //如果没有找到开始，则不找结束点
                    //再找到数值的结束点
                    if(c!='.' && (c<'0' || c>'9'))
                        return value.Substring(valueStart, start-valueStart);
                }
            }
            //如果到最后也没有结尾符，直接取字符串到末尾
            //最后就一个数值时会直接退出循环，也相当于没找到结束符
            if(valueStart >=0 && start == value.Length)
                return value.Substring(valueStart, value.Length-valueStart);

            return string.Empty;
        }

        private static readonly List<int> workingIntlist = new List<int>(8);
        private static bool TryParseIntArray(string value, out int[] array)
        {
            workingIntlist.Clear();
            for(int i = 0; i<value.Length; i++)
            {
                char c = value[i];
                if('0' <= c && c <= '9')
                {
                    //发现数值，开始读取
                    int n = 0;
                    while('0' <= c && c <= '9')
                    {
                        n = n*10 + c - '0';
                        i++;
                        if(i >= value.Length) break;
                        c = value[i];
                    }
                    workingIntlist.Add(n);
                }
            }
            array = workingIntlist.ToArray();
            if(array.Length == 0)
                return false;
            return true;
        }

        private static string GetConfScriptName(string sheetName)
        {
            return $"Conf{sheetName}";
        }

        private static void PressAnyKeyToExist()
        {
            Console.WriteLine("Press any to exist");
            Console.ReadKey();
        }
    }

    struct Vector2
    {
        public float x;
        public float y;

        public Vector2(float x, float y)
        {
            this.x = x;
            this.y = y;
        }

        public void WriteXml(XmlWriter write)
        {
            write.WriteStartElement("x"); write.WriteValue(x); write.WriteEndElement();
            write.WriteStartElement("y"); write.WriteValue(y); write.WriteEndElement();
        }
    }

    struct Vector3
    {
        public float x;
        public float y;
        public float z;

        public Vector3(float x, float y, float z)
        {
            this.x = x;
            this.y = y;
            this.z = z;
        }

        public void WriteXml(XmlWriter write)
        {
            write.WriteStartElement("x"); write.WriteValue(x); write.WriteEndElement();
            write.WriteStartElement("y"); write.WriteValue(y); write.WriteEndElement();
            write.WriteStartElement("z"); write.WriteValue(z); write.WriteEndElement();
        }
    }
}
