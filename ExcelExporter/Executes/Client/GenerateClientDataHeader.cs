using System.IO;
using System;
using log4net;

namespace ExcelExporter
{
    public class GenerateClientDataHeader : IDisposable
    {
        protected static readonly ILog Log = LogManager.GetLogger(typeof(GenerateClientDataHeader));

        public GenerateClientDataHeader(string dataHeaderPath)
        {
            DataHeaderPath = dataHeaderPath;
            DataHeaderFileName = Utils.ClientDataFileName;
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

                    foreach (System.Data.DataTable rawDataTable in rawDataSet.Tables)
                    {
                        if (!Utils.IsDataTable(rawDataTable))
                            continue;

                        string tableName = rawDataTable.TableName;

                        textWriter.WriteLine();
                        textWriter.WriteLine("USTRUCT(BlueprintType)");
                        textWriter.WriteLine(
                            string.Format("struct F{0} : public FTableRowBase", tableName));
                        textWriter.WriteLine("{");
                        textWriter.WriteLine("    GENERATED_USTRUCT_BODY()");
                        textWriter.WriteLine();
                        textWriter.WriteLine("public:");

                        int columnCount = rawDataTable.Columns.Count;
                        for (int i = 0; i < columnCount; ++i)
                        {
                            string columnName = rawDataTable.Columns[i].ColumnName;
                            string typeValue = rawDataTable.Rows[0].ItemArray[i].ToString();

                            if (!Utils.IsRemovableClientColumn(columnName, typeValue))
                            {
                                string typeName = Utils.GetClientTypeName(typeValue);

                                if (i != 0)
                                {
                                    textWriter.WriteLine();
                                }

                                textWriter.WriteLine("    UPROPERTY(EditAnywhere, BlueprintReadWrite, Category = RowData)");
                                textWriter.WriteLine(string.Format("    {0} {1};",
                                    typeName, columnName));
                            }
                        }

                        textWriter.WriteLine("};");
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

        public string DataHeaderPath { get; set; }
        public string DataHeaderFileName { get; set; }
    }
}
