using System.IO;
using System;
using log4net;

namespace ExcelExporter
{
    public class GenerateServerDataHeader : IDisposable
    {
        protected static readonly ILog Log = LogManager.GetLogger(typeof(GenerateServerDataHeader));

        public GenerateServerDataHeader(string dataHeaderPath)
        {
            DataHeaderPath = dataHeaderPath;
            DataHeaderFileName = Utils.ServerDataFileName;
        }

        public bool Execute(System.Data.DataSet rawDataSet)
        {
            try
            {
                Directory.CreateDirectory(DataHeaderPath);
                string dataHeaderFilePath = Path.Combine(DataHeaderPath, DataHeaderFileName);

                using (TextWriter textWriter =
                    new StreamWriter(dataHeaderFilePath, false, System.Text.Encoding.UTF8))
                {
                    textWriter.WriteLine("//////////////////////////////////");
                    textWriter.WriteLine("// This file is auto generated. //");
                    textWriter.WriteLine("//////////////////////////////////");
                    textWriter.WriteLine("#pragma once");
                    textWriter.WriteLine(string.Format("#include \"{0}\"", Utils.ServerEnumFileName));
                    textWriter.WriteLine();
                    textWriter.WriteLine(Utils.ServerGamedataNamespaceBegin);
                    textWriter.WriteLine();
                    textWriter.WriteLine("struct GameDataRow");
                    textWriter.WriteLine("{");
                    textWriter.WriteLine("    virtual ~GameDataRow() = default;");
                    textWriter.WriteLine("};");
                    textWriter.WriteLine();

                    foreach (System.Data.DataTable rawDataTable in rawDataSet.Tables)
                    {
                        if (!Utils.IsDataTable(rawDataTable))
                            continue;

                        string tableName = rawDataTable.TableName;

                        textWriter.WriteLine(
                            string.Format("struct {0} : public GameDataRow", tableName));
                        textWriter.WriteLine("{");

                        {
                            // Data members
                            int columnCount = rawDataTable.Columns.Count;
                            for (int i = 0; i < columnCount; ++i)
                            {
                                string columnName = rawDataTable.Columns[i].ColumnName;
                                string typeValue = rawDataTable.Rows[0].ItemArray[i].ToString();

                                if (!Utils.IsRemovableServerColumn(columnName, typeValue))
                                {
                                    string typeName = Utils.GetServerTypeName(typeValue);

                                    if (Utils.GetTypeColumn(rawDataTable, typeName, columnName,
                                        out Tuple<string, string> typeColumns))
                                    {
                                        if (Utils.IsEnumType(typeName))
                                        {
                                            textWriter.WriteLine(string.Format(
                                                "    {0} {1} = {0}::_from_integral(0);",
                                                typeColumns.Item1, typeColumns.Item2));
                                        }
                                        else
                                        {
                                            textWriter.WriteLine(string.Format("    {0} {1};",
                                                typeColumns.Item1, typeColumns.Item2));
                                        }
                                    }
                                }
                            }

                            textWriter.WriteLine();
                        }

                        {
                            // serialize function
                            textWriter.WriteLine("    template <typename Archive>");
                            textWriter.WriteLine("    void serialize(Archive& archive)");
                            textWriter.WriteLine("    {");
                            {
                                int columnCount = rawDataTable.Columns.Count;
                                for (int i = 0; i < columnCount; ++i)
                                {
                                    string columnName = rawDataTable.Columns[i].ColumnName;
                                    string typeValue = rawDataTable.Rows[0].ItemArray[i].ToString();

                                    if (!Utils.IsRemovableServerColumn(columnName, typeValue))
                                    {
                                        string tempTypeName = Utils.GetServerTempTypeName(typeValue);
                                        string tempColumnName = "temp_" + columnName;

                                        textWriter.WriteLine(string.Format("        {0} {1};",
                                            tempTypeName, tempColumnName));
                                    }
                                }
                            }
                            textWriter.WriteLine();
                            textWriter.WriteLine("        if (typeid(archive) == typeid(cereal::JSONOutputArchive))");
                            textWriter.WriteLine("        {");
                            {
                                int columnCount = rawDataTable.Columns.Count;
                                for (int i = 0; i < columnCount; ++i)
                                {
                                    string columnName = rawDataTable.Columns[i].ColumnName;
                                    string typeValue = rawDataTable.Rows[0].ItemArray[i].ToString();

                                    if (!Utils.IsRemovableServerColumn(columnName, typeValue))
                                    {
                                        string typeName = Utils.GetServerTypeName(typeValue);
                                        string tempColumnName = "temp_" + columnName;

                                        if (Utils.IsEnumType(typeName))
                                        {
                                            textWriter.WriteLine(string.Format(
                                                "            {0} = {1}._to_string();",
                                                tempColumnName, columnName));
                                        }
                                        else if (Utils.IsArrayType(columnName,
                                            out string arrayName, out int index))
                                        {
                                            if (Utils.IsVectorOrRotatorType(typeName))
                                            {
                                                textWriter.WriteLine(string.Format(
                                                    "            {0} = {1}[{2}].ToString();",
                                                    tempColumnName, arrayName, index));
                                            }
                                            else
                                            {
                                                textWriter.WriteLine(string.Format(
                                                    "            {0} = {1}[{2}];",
                                                    tempColumnName, arrayName, index));
                                            }
                                        }
                                        else if (Utils.IsVectorOrRotatorType(typeName))
                                        {
                                            textWriter.WriteLine(string.Format(
                                                "            {0} = {1}.ToString();",
                                                tempColumnName, columnName));
                                        }
                                        else
                                        {
                                            textWriter.WriteLine(string.Format(
                                                "            {0} = {1};",
                                                tempColumnName, columnName));
                                        }
                                    }
                                }
                            }
                            textWriter.WriteLine("        }");
                            textWriter.WriteLine();
                            textWriter.WriteLine("        archive(");

                            {
                                int columnCount = rawDataTable.Columns.Count;
                                for (int i = 0; i < columnCount; ++i)
                                {
                                    string columnName = rawDataTable.Columns[i].ColumnName;
                                    string typeValue = rawDataTable.Rows[0].ItemArray[i].ToString();

                                    if (!Utils.IsRemovableServerColumn(columnName, typeValue))
                                    {
                                        string tempColumnName = "temp_" + columnName;

                                        if (i == 0)
                                        {
                                            textWriter.WriteLine(string.Format(
                                                "              ::cereal::make_nvp(\"{0}\", {1})",
                                                columnName, tempColumnName));
                                        }
                                        else
                                        {
                                            textWriter.WriteLine(string.Format(
                                                "            , ::cereal::make_nvp(\"{0}\", {1})",
                                                columnName, tempColumnName));
                                        }
                                    }
                                }
                            }

                            textWriter.WriteLine("        );");
                            textWriter.WriteLine();
                            textWriter.WriteLine("        if (typeid(archive) == typeid(cereal::JSONInputArchive))");
                            textWriter.WriteLine("        {");
                            {

                                int columnCount = rawDataTable.Columns.Count;
                                for (int i = 0; i < columnCount; ++i)
                                {
                                    string columnName = rawDataTable.Columns[i].ColumnName;
                                    string typeValue = rawDataTable.Rows[0].ItemArray[i].ToString();

                                    if (!Utils.IsRemovableServerColumn(columnName, typeValue))
                                    {
                                        string typeName = Utils.GetServerTypeName(typeValue);
                                        string tempColumnName = "temp_" + columnName;

                                        if (Utils.IsEnumType(typeName))
                                        {
                                            textWriter.WriteLine(string.Format(
                                                "            {0} = {1}::_from_string({2}.c_str());",
                                                columnName, typeName, tempColumnName));
                                        }
                                        else if (Utils.IsArrayType(columnName,
                                            out string arrayName, out int index))
                                        {
                                            if (Utils.IsVectorOrRotatorType(typeName))
                                            {
                                                textWriter.WriteLine(string.Format(
                                                    "            {0}[{1}].InitFromString({2});",
                                                    arrayName, index, tempColumnName));
                                            }
                                            else
                                            {
                                                textWriter.WriteLine(string.Format(
                                                    "            {0}[{1}] = {2};",
                                                    arrayName, index, tempColumnName));
                                            }
                                        }
                                        else if (Utils.IsVectorOrRotatorType(typeName))
                                        {
                                            textWriter.WriteLine(string.Format(
                                                "            {0}.InitFromString({1});",
                                                columnName, tempColumnName));
                                        }
                                        else
                                        {
                                            textWriter.WriteLine(string.Format(
                                                "            {0} = {1};",
                                                columnName, tempColumnName));
                                        }
                                    }
                                }
                            }
                            textWriter.WriteLine("        }");
                            textWriter.WriteLine("    }");
                        }

                        textWriter.WriteLine("};");
                        textWriter.WriteLine();

                        {
                            // operator <<
                            textWriter.WriteLine("template <typename traits>");
                            textWriter.WriteLine("inline std::basic_ostream<wchar_t, traits>&");
                            textWriter.WriteLine(string.Format("operator<< (std::basic_ostream<wchar_t, traits>& os, const {0}& rhs)", tableName));
                            textWriter.WriteLine("{");
                            textWriter.WriteLine("    std::stringstream strStream;");
                            textWriter.WriteLine();
                            textWriter.WriteLine("    {");
                            textWriter.WriteLine("        cereal::JSONOutputArchive archive(strStream, cereal::JSONOutputArchive::Options::NoIndent());");
                            textWriter.WriteLine("        auto rowData(rhs);");
                            textWriter.WriteLine("        rowData.serialize(archive);");
                            textWriter.WriteLine("    }");
                            textWriter.WriteLine();
                            textWriter.WriteLine("    return os << FromUTF8(boost::replace_all_copy(strStream.str(), \"\\n\", \"\"));");
                            textWriter.WriteLine("}");
                            textWriter.WriteLine();
                        }
                    }

                    textWriter.WriteLine(Utils.ServerGamedataNamespaceEnd);
                    textWriter.Close();
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex);
                return false;
            }

            return true;
        }

        public void Dispose()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public string DataHeaderPath { get; set; }
        public string DataHeaderFileName { get; set; }
    }
}
