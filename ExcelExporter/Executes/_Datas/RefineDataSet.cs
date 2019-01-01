using System.Collections.Generic;
using System;
using log4net;

namespace ExcelExporter
{
    class RefineDataSet : IDisposable
    {
        protected static readonly ILog Log = LogManager.GetLogger(typeof(RefineDataSet));

        public RefineDataSet()
        {
        }

        public bool ExecuteServerData(System.Data.DataSet rawDataSet,
            ref System.Data.DataSet refinedServerDataSet)
        {
            try
            {
                foreach (System.Data.DataTable rawDataTable in rawDataSet.Tables)
                {
                    if (!Utils.IsDataTable(rawDataTable))
                        continue;

                    string tableName = rawDataTable.TableName;
                    System.Data.DataTable dataTable = new System.Data.DataTable(tableName);

                    // Remove colums
                    List<string> removableColumnNames = new List<string>();

                    // Add Column
                    int columnCount = rawDataTable.Columns.Count;
                    for (int i = 0; i < columnCount; ++i)
                    {
                        string columnName = rawDataTable.Columns[i].ColumnName;
                        string typeValue = rawDataTable.Rows[0].ItemArray[i].ToString();

                        if (columnName.StartsWith("#") || typeValue.StartsWith("#"))
                        {
                            typeValue = "string";
                        }

                        dataTable.Columns.Add(
                            new System.Data.DataColumn(columnName, GetType(typeValue)));

                        if (Utils.IsRemovableServerColumn(columnName, typeValue))
                        {
                            removableColumnNames.Add(columnName);
                        }
                    }

                    // Add Rows
                    int rowCount = rawDataTable.Rows.Count;
                    for (int i = 0; i < rowCount; ++i)
                    {
                        if (i == 0)
                            continue;

                        if (!IsPassedNullableCheck(rawDataTable, i))
                        {
                            Log.Error("Fail to pass nullable check.");
                            return false;
                        }

                        System.Data.DataRow dataRow = dataTable.NewRow();
                        dataRow.ItemArray = GetItemArrayFillDefaultValue(rawDataTable, i);
                        dataTable.Rows.Add(dataRow);
                    }

                    // Remove columns
                    foreach (string removeColumnName in removableColumnNames)
                    {
                        dataTable.Columns.Remove(removeColumnName);
                    }

                    refinedServerDataSet.Tables.Add(dataTable);
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex);
                return false;
            }

            return true;
        }

        public bool ExecuteClientData(System.Data.DataSet rawDataSet,
            ref System.Data.DataSet refinedClientDataSet)
        {
            try
            {
                foreach (System.Data.DataTable rawDataTable in rawDataSet.Tables)
                {
                    if (!Utils.IsDataTable(rawDataTable))
                        continue;

                    string tableName = rawDataTable.TableName;
                    System.Data.DataTable dataTable = new System.Data.DataTable(tableName);

                    // Remove colums
                    List<string> removableColumnNames = new List<string>();

                    // Add Column
                    int columnCount = rawDataTable.Columns.Count;
                    for (int i = 0; i < columnCount; ++i)
                    {
                        string columnName = rawDataTable.Columns[i].ColumnName;
                        string typeValue = rawDataTable.Rows[0].ItemArray[i].ToString();

                        if (columnName.StartsWith("#") || typeValue.StartsWith("#"))
                        {
                            typeValue = "string";
                        }

                        dataTable.Columns.Add(
                            new System.Data.DataColumn(columnName, GetType(typeValue)));

                        if (Utils.IsRemovableClientColumn(columnName, typeValue))
                        {
                            removableColumnNames.Add(columnName);
                        }
                    }

                    // Add Rows
                    int rowCount = rawDataTable.Rows.Count;
                    for (int i = 0; i < rowCount; ++i)
                    {
                        if (i == 0)
                            continue;

                        if (!IsPassedNullableCheck(rawDataTable, i))
                        {
                            Log.Error("Fail to pass nullable check.");
                            return false;
                        }

                        System.Data.DataRow dataRow = dataTable.NewRow();
                        dataRow.ItemArray = GetItemArrayFillDefaultValue(rawDataTable, i);
                        dataTable.Rows.Add(dataRow);
                    }

                    // Remove columns
                    foreach (string removeColumnName in removableColumnNames)
                    {
                        dataTable.Columns.Remove(removeColumnName); 
                    }

                    refinedClientDataSet.Tables.Add(dataTable);
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

        private Type GetType(string typeValue)
        {
            typeValue = typeValue.TrimEnd('_', 'c', 's', '?');

            switch (typeValue)
            {
                case "int32": return typeof(int);
                case "int64": return typeof(long);
                case "float": return typeof(float);
                case "double": return typeof(double);
            }

            return typeof(string);
        }

        private string GetDefaultValue(string typeValue)
        {
            typeValue = typeValue.TrimEnd('_', 'c', 's', '?');

            switch (typeValue)
            {
                case "int32": return "0";
                case "int64": return "0";
                case "float": return "0.0";
                case "double": return "0.0";
                case "vector": return "(X=0,Y=0,Z=0)";
                case "rotator": return "(P=0,Y=0,R=0)";
            }

            return "";
        }

        private bool IsPassedNullableCheck(System.Data.DataTable rawDataTable, int rowIndex)
        {
            int columnCount = rawDataTable.Columns.Count;
            for (int i = 0; i < columnCount; ++i)
            {
                string columnName = rawDataTable.Columns[i].ColumnName;
                string typeValue = rawDataTable.Rows[0].ItemArray[i].ToString();
                bool isNullable = typeValue.Contains("?");
                string cell = rawDataTable.Rows[rowIndex].ItemArray[i].ToString();

                if (columnName.StartsWith("#") || typeValue.StartsWith("#"))
                    return true;

                if (!isNullable && cell.Length == 0)
                {
                    Log.ErrorFormat("Fail to pass null. table: {0}, Id: {1}, type: {2}, column: {3}",
                        rawDataTable.TableName,
                        rawDataTable.Rows[rowIndex].ItemArray[0].ToString(),
                        typeValue,
                        rawDataTable.Columns[i].ColumnName);

                    return false;
                }
            }

            return true;
        }

        private object[] GetItemArrayFillDefaultValue(
            System.Data.DataTable rawDataTable, int rowIndex)
        {
            int columnCount = rawDataTable.Columns.Count;

            object[] itemArray = new object[columnCount];
            for (int i = 0; i < columnCount; ++i)
            {
                string columnName = rawDataTable.Columns[i].ColumnName;
                string typeValue = rawDataTable.Rows[0].ItemArray[i].ToString();
                bool isNullable = typeValue.Contains("?");
                string cell = rawDataTable.Rows[rowIndex].ItemArray[i].ToString();

                if (columnName.StartsWith("#") || typeValue.StartsWith("#"))
                {
                    itemArray[i] = cell;
                }
                else if (isNullable && cell.Length == 0)
                {
                    itemArray[i] = GetDefaultValue(typeValue);
                }
                else
                {
                    itemArray[i] = cell;
                }
            }

            return itemArray;
        }
    }
}
