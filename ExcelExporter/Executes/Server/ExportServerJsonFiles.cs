using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using System.IO;
using System;
using log4net;

namespace ExcelExporter
{
    public class ExportServerJsonFiles : IDisposable
    {
        protected static readonly ILog Log = LogManager.GetLogger(typeof(ExportServerJsonFiles));

        public ExportServerJsonFiles(string exportPath)
        {
            JsonExportPath = exportPath;
        }

        public bool Execute(System.Data.DataSet refinedDataSet, System.Data.DataSet rawDataSet)
        {
            try
            {
                Directory.CreateDirectory(JsonExportPath);

                foreach (System.Data.DataTable refinedDataTable in refinedDataSet.Tables)
                {
                    if (!Utils.IsDataTable(refinedDataTable))
                        continue;

                    string tableName = refinedDataTable.TableName;
                    int dataHashCode = 0;
                    foreach (System.Data.DataRow row in refinedDataTable.Rows)
                    {
                        foreach (var item in row.ItemArray)
                        {
                            dataHashCode += item.GetHashCode();
                        }
                    }

                    JObject jHeaderObject = new JObject();

                    System.Data.DataTable rawDataTable = rawDataSet.Tables[tableName];
                    if (rawDataTable != null)
                    {
                        int columnCount = rawDataTable.Columns.Count;
                        for (int i = 0; i < columnCount; ++i)
                        {
                            string columnName = rawDataTable.Columns[i].ColumnName;
                            string typeValue = rawDataTable.Rows[0].ItemArray[i].ToString();

                            if (!Utils.IsRemovableServerColumn(columnName, typeValue))
                            {
                                string typeName = Utils.GetServerTypeName(typeValue);
                                jHeaderObject.Add(columnName, typeName);
                            }
                        }
                    }

                    JObject jExportObject = new JObject();
                    jExportObject.Add("infos", new JObject
                    {
                        { "rowcount", refinedDataTable.Rows.Count },
                        { "dataHashCode", dataHashCode },
                    });
                    jExportObject.Add("header", jHeaderObject);
                    jExportObject.Add("rows", JToken.FromObject(refinedDataTable));

                    string fileName = tableName + ".json";
                    string jsonFilePath = Path.Combine(JsonExportPath, fileName);
                    using (StreamWriter writer = new StreamWriter(jsonFilePath, false,
                        System.Text.Encoding.UTF8))
                    {
                        string exportJson = JsonConvert.SerializeObject(jExportObject,
                            Formatting.Indented);

                        writer.Write(exportJson);
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

        public string JsonExportPath { get; set; }
    }
}
