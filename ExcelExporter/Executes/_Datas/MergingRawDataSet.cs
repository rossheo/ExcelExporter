using DotNet.Collections.Generic;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System;
using log4net;

namespace ExcelExporter
{
    class MergingRawDataSet : IDisposable
    {
        protected static readonly ILog Log = LogManager.GetLogger(typeof(MergingRawDataSet));

        public MergingRawDataSet()
        {
        }

        public bool Execute(ref System.Data.DataSet rawDataSet)
        {
            using (System.Data.DataSet mergedRawDataSet = new System.Data.DataSet())
            {
                MultiMapList<string, string> multimapList = new MultiMapList<string, string>();

                foreach (System.Data.DataTable rawDataTable in rawDataSet.Tables)
                {
                    string tableName = rawDataTable.TableName;

                    if (tableName.Contains("_"))
                    {
                        string[] splitedTableName = tableName.Split('_');

                        if (splitedTableName.Length != 2)
                        {
                            Log.ErrorFormat("Error. Table's name has twice underbar('_'). {0}",
                                tableName);
                            Trace.Assert(false);
                            return false;
                        }

                        if (!multimapList.TryToAddMapping(splitedTableName[0], tableName))
                        {
                            Log.ErrorFormat("Fail to add. {0}, {1}",
                                splitedTableName[0], tableName);
                            return false;
                        }
                    }
                    else
                    {
                        if (!multimapList.TryToAddMapping(tableName, tableName))
                        {
                            Log.ErrorFormat("Fail to add. {0}, {1}",
                                tableName, tableName);
                            return false;
                        }
                    }
                }

                foreach (var kvp in multimapList)
                {
                    string key = kvp.Key;
                    List<string> values = kvp.Value;

                    foreach (string value in values)
                    {
                        string tableName = value;
                        System.Data.DataTable rawDataTable = rawDataSet.Tables[tableName];
                        if (rawDataTable != null)
                        {
                            System.Data.DataTable addOrMergeTable = rawDataTable.Copy();
                            addOrMergeTable.TableName = key;

                            if (!mergedRawDataSet.Tables.Contains(key))
                            {
                                mergedRawDataSet.Tables.Add(addOrMergeTable);
                            }
                            else
                            {
                                var compareSourceRow =
                                    mergedRawDataSet.Tables[key].Rows[0].ItemArray
                                    .Aggregate((lhs, rhs) => lhs + ", " + rhs);
                                var compareDestRow = addOrMergeTable.Rows[0].ItemArray
                                    .Aggregate((lhs, rhs) => lhs + ", " + rhs);

                                if (!compareSourceRow.Equals(compareDestRow))
                                {
                                    Log.ErrorFormat("Table Schema is not matched." +
                                        " {0}, source: {1}, dest: {2}",
                                        key, compareSourceRow, compareDestRow);
                                    return false;
                                }

                                addOrMergeTable.Rows.RemoveAt(0);
                                mergedRawDataSet.Tables[key].Merge(addOrMergeTable);
                            }
                        }
                    }
                }

                rawDataSet = mergedRawDataSet;
            }

            return true;
        }

        public void Dispose()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
