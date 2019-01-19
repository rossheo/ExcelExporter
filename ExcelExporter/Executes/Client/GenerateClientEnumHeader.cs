using System.IO;
using System;
using log4net;

namespace ExcelExporter
{
    public class GenerateClientEnumHeader : IDisposable
    {
        protected static readonly ILog Log = LogManager.GetLogger(typeof(GenerateClientEnumHeader));

        public GenerateClientEnumHeader(string enumHeaderPath)
        {
            EnumHeaderPath = enumHeaderPath;
            EnumHeaderFileName = Utils.ClientEnumFileName;
        }

        public bool Execute(System.Data.DataSet rawDataSet)
        {
            try
            {
                Directory.CreateDirectory(EnumHeaderPath);
                string enumHeaderFilePath = Path.Combine(EnumHeaderPath, EnumHeaderFileName);

                using (TextWriter textWriter =
                    new StreamWriter(enumHeaderFilePath, false, System.Text.Encoding.UTF8))
                {
                    textWriter.WriteLine("//////////////////////////////////");
                    textWriter.WriteLine("// This file is auto generated. //");
                    textWriter.WriteLine("//////////////////////////////////");
                    textWriter.WriteLine("#pragma once");
                    textWriter.WriteLine();

                    foreach (System.Data.DataTable rawDataTable in rawDataSet.Tables)
                    {
                        if (!Utils.IsEnumTable(rawDataTable))
                            continue;

                        string tableName = rawDataTable.TableName;
                        string enumName = Utils.GetEnumName(tableName);

                        // Comments
                        textWriter.WriteLine("/*");
                        foreach (System.Data.DataRow row in rawDataTable.Rows)
                        {
                            string firstColumn = row.ItemArray[0].ToString();
                            string thirdColumn = row.ItemArray[2].ToString();

                            if (thirdColumn.Length > 0)
                            {
                                textWriter.WriteLine(
                                    string.Format("{0} : {1}", firstColumn, thirdColumn));
                            }
                        }
                        textWriter.WriteLine("*/");

                        // Enum
                        textWriter.WriteLine("UENUM(BlueprintType)");
                        textWriter.WriteLine(string.Format("enum class {0} : uint8", enumName));
                        textWriter.WriteLine("{");
                        foreach (System.Data.DataRow row in rawDataTable.Rows)
                        {
                            string firstColumn = row.ItemArray[0].ToString();
                            string secondColumn = row.ItemArray[1].ToString();

                            if (firstColumn.StartsWith("#"))
                                continue;

                            if (secondColumn.Length > 0)
                            {
                                textWriter.WriteLine(
                                    string.Format("    {0} = {1, -18} UMETA(DisplayName = \"{0}\"),",
                                    firstColumn, secondColumn));
                            }
                            else
                            {
                                textWriter.WriteLine(
                                    string.Format("    {0, -25} UMETA(DisplayName = \"{0}\"),",
                                    firstColumn));
                            }
                        }
                        textWriter.WriteLine("};");
                        textWriter.WriteLine();
                    }
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

        public string EnumHeaderPath { get; set; }
        public string EnumHeaderFileName { get; set; }
    }
}
