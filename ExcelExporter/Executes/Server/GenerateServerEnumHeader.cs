using System.IO;
using System;
using log4net;

namespace ExcelExporter
{
    public class GenerateServerEnumHeader : IDisposable
    {
        protected static readonly ILog Log = LogManager.GetLogger(typeof(GenerateServerEnumHeader));

        public GenerateServerEnumHeader(string enumHeaderPath)
        {
            EnumHeaderPath = enumHeaderPath;
            EnumHeaderFileName = "gamedata_enum.h";
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
                    textWriter.WriteLine(Utils.GetServerNameSpaceBegin());
                    textWriter.WriteLine();

                    foreach (System.Data.DataTable rawDataTable in rawDataSet.Tables)
                    {
                        if (!Utils.IsEnumTable(rawDataTable))
                            continue;

                        string tableName = rawDataTable.TableName;

                        textWriter.Write(string.Format("BETTER_ENUM({0}{1}, uint32, ", Utils.EnumPrefix, tableName));

                        string enumStrings = string.Empty;
                        foreach (System.Data.DataRow row in rawDataTable.Rows)
                        {
                            if (enumStrings.Length != 0)
                            {
                                enumStrings += ", ";
                            }

                            string firstColumn = row.ItemArray[0].ToString();
                            string secondColumn = row.ItemArray[1].ToString();

                            if (firstColumn.StartsWith("#"))
                                continue;

                            if (secondColumn.Length == 0)
                            {
                                enumStrings += string.Format("{0}", firstColumn);
                            }
                            else
                            {
                                enumStrings += string.Format("{0}={1}", firstColumn, secondColumn);
                            }
                        }

                        textWriter.Write(enumStrings);
                        textWriter.WriteLine(");");
                    }

                    textWriter.WriteLine();
                    textWriter.WriteLine(Utils.GetServerNameSpaceEnd());
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
