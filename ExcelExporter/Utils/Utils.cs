using DotNet.Collections.Generic;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text.RegularExpressions;

namespace ExcelExporter
{
    public class Utils
    {
        public static string EnumPrefix { get { return "Enum_"; } }

        public static string GetServerNameSpaceBegin()
        {
            return "namespace rh::gamedata\r\n{";
        }

        public static string GetServerNameSpaceEnd()
        {
            return "} // namespace rh::gamedata";
        }

        public static bool IsDataTable(System.Data.DataTable rawDataTable)
        {
            if (rawDataTable.TableName.StartsWith("#"))
                return false;

            return rawDataTable.Columns[0].ColumnName == "Id";
        }

        public static bool IsEnumTable(System.Data.DataTable rawDataTable)
        {
            if (rawDataTable.TableName.StartsWith("#"))
                return false;

            return (rawDataTable.Columns[0].ColumnName == "Enum")
                && (rawDataTable.Columns[1].ColumnName == "Value")
                && (rawDataTable.Columns[2].ColumnName == "Description");
        }

        public static bool IsEnumType(string dataType)
        {
            return dataType.StartsWith(EnumPrefix);
        }

        public static bool IsArrayType(string dataType, out string arrayName, out int index)
        {
            Regex regex = new Regex("(\\w+)_(\\d+)");
            Match m = regex.Match(dataType);
            if (!m.Success)
            {
                arrayName = string.Empty;
                index = 0;
                return false;
            }

            arrayName = m.Groups[1].Value + "s";
            index = Convert.ToInt32(m.Groups[2].Value) - 1;
            return true;
        }

        public static string GetEnumName(string tableName)
        {
            return string.Format("{0}{1}", EnumPrefix, tableName);
        }

        public static string GetServerTypeName(string dataType)
        {
            string trimedDataType = dataType.TrimEnd('_', 'c', 's', '?');

            switch (trimedDataType)
            {
                case "int32": return "int32";
                case "int64": return "int64";
                case "float": return "float";
                case "double": return "double";
                case "vector": return "FVector";
                case "rotator": return "FRotator";
                case "string": return "std::string";
            }

            return trimedDataType;
        }

        public static string GetServerTempTypeName(string dataType)
        {
            string trimedDataType = dataType.TrimEnd('_', 'c', 's', '?');

            switch (trimedDataType)
            {
                case "int32": return "int32";
                case "int64": return "int64";
                case "float": return "float";
                case "double": return "double";
            }

            return "std::string";
        }

        public static string GetClientTypeName(string dataType)
        {
            string trimedDataType = dataType.TrimEnd('_', 'c', 's', '?');

            switch (trimedDataType)
            {
                case "int32": return "int32";
                case "int64": return "int64";
                case "float": return "float";
                case "double": return "double";
                case "vector": return "FVector";
                case "rotator": return "FRotator";
                case "string": return "FString";
            }

            return trimedDataType;
        }

        public static bool IsRemovableServerColumn(string columnName, string typeValue)
        {
            if (columnName.StartsWith("#"))
                return true;

            if (typeValue.StartsWith("#"))
                return true;

            if (!typeValue.Contains("_"))
            {
                Trace.Assert(false);
                return true;
            }

            string[] splitedTypeValues = typeValue.Split('_');

            string lastSplitedTypeValue = splitedTypeValues[splitedTypeValues.Length - 1];
            if (!lastSplitedTypeValue.Contains("s"))
                return true;

            return false;
        }

        public static bool IsRemovableClientColumn(string columnName, string typeValue)
        {
            if (columnName.StartsWith("#"))
                return true;

            if (typeValue.StartsWith("#"))
                return true;

            if (!typeValue.Contains("_"))
            {
                Trace.Assert(false);
                return true;
            }

            string[] splitedTypeValues = typeValue.Split('_');

            string lastSplitedTypeValue = splitedTypeValues[splitedTypeValues.Length - 1];
            if (!lastSplitedTypeValue.Contains("c"))
                return true;

            return false;
        }

        public static bool GetTypeColumn(System.Data.DataTable rawDataTable,
            string typeName, string columnName, out Tuple<string, string> typeColumns)
        {
            string columnKey = string.Empty;
            string columnValue = string.Empty;

            {
                Regex regex = new Regex("(\\w+)_(\\d+)");
                Match m = regex.Match(columnName);
                if (!m.Success)
                {
                    typeColumns = new Tuple<string, string>(typeName, columnName);
                    return true;
                }

                columnKey = m.Groups[1].Value;
                columnValue = m.Groups[2].Value;

                int columnValueToInt = Convert.ToInt32(columnValue);
                if (columnValueToInt > 1)
                {
                    Trace.Assert(columnValueToInt != 0);

                    typeColumns = new Tuple<string, string>(string.Empty, string.Empty);
                    return false;
                }
            }

            MultiMapList<string, string> multiMapList = new MultiMapList<string, string>();

            {
                int columnCount = rawDataTable.Columns.Count;
                for (int i = 0; i < columnCount; ++i)
                {
                    string rawDataColumnName = rawDataTable.Columns[i].ColumnName;

                    Regex regex = new Regex("(\\w+)_(\\d+)");
                    Match m = regex.Match(rawDataColumnName);
                    if (m.Success)
                    {
                        string key = m.Groups[1].Value;
                        string value = m.Groups[2].Value;

                        multiMapList.TryToAddMapping(key, value);
                    }
                }
            }

            List<string> columnValues = new List<string>();
            if (!multiMapList.TryGetValue(columnKey, out columnValues))
            {
                typeColumns = new Tuple<string, string>(string.Empty, string.Empty);
                return false;
            }

            int columnKeyCount = columnValues.Count;
            string arrayTypeName = string.Format("std::array<{0}, {1}>", typeName, columnKeyCount);

            typeColumns = new Tuple<string, string>(arrayTypeName, columnKey + "s");
            return true;
        }
    }
}
