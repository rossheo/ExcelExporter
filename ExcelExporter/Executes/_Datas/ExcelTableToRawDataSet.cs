using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System;
using log4net;

namespace ExcelExporter
{
    class ExcelTableToRawDataSet : IDisposable
    {
        protected static readonly ILog Log = LogManager.GetLogger(typeof(ExcelTableToRawDataSet));

        public struct SYSTEMTIME
        {
            public ushort wYear;
            public ushort wMonth;
            public ushort wDayOfWeek;
            public ushort wDay;
            public ushort wHour;
            public ushort wMinute;
            public ushort wSecond;
            public ushort wMilliseconds;
        }

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern uint SetSystemTime(ref SYSTEMTIME systemTime);

        public ExcelTableToRawDataSet(string excelPath, bool isFixedDate)
        {
            ExcelPath = excelPath;
            IsFixedDate = isFixedDate;

            DateTime utcNow = DateTime.UtcNow;
            CurrentYear = utcNow.Year;
            CurrentMonth = utcNow.Month;
            CurrentDay = utcNow.Day;
        }

        public bool Execute(ref System.Data.DataSet rawDataSet)
        {
            Application excelApplication = new Application();

            List<Workbook> workBooks = new List<Workbook>();

            try
            {
                SetFixedSystemTime();

                if (File.Exists(ExcelPath))
                {
                    Workbook workbook = OpenExcelWorkbook(excelApplication, ExcelPath);
                    if (workbook != null)
                    {
                        workBooks.Add(workbook);
                    }
                }
                else
                {
                    string[] excelFilePaths = Directory.GetFiles(ExcelPath, "*.xlsx");

                    foreach (string excelFilePath in excelFilePaths)
                    {
                        string excelFileName = Path.GetFileName(excelFilePath);
                        if (excelFileName.StartsWith("~") || excelFileName.StartsWith("#"))
                            continue;

                        Workbook workbook = OpenExcelWorkbook(excelApplication, excelFilePath);
                        if (workbook != null)
                        {
                            workBooks.Add(workbook);
                        }
                    }
                }

                foreach (Workbook workbook in workBooks)
                {
                    List<Worksheet> workSheets = GetWorkSheets(workbook);
                    if (workSheets == null)
                        return false;

                    List<ListObject> listObjects = GetListObjects(workSheets);
                    if (listObjects == null)
                        return false;

                    ListObjectsToRawDataTable(listObjects, ref rawDataSet);
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex.ToString());
            }
            finally
            {
                SetCurrentSystemTime();

                foreach (Workbook workbook in workBooks)
                {
                    workbook.Close(Type.Missing, Type.Missing, Type.Missing);
                    Release(workbook);
                }

                excelApplication.Quit();
                Release(excelApplication);
            }

            return true;
        }

        public void Dispose()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private Workbook OpenExcelWorkbook(Application excelApplication, string excelPath)
        {
            FileInfo excelFileInfo = new FileInfo(excelPath);
            if (!excelFileInfo.Exists)
            {
                Log.WarnFormat("File is not exist. {0}", excelPath);
                return null;
            }

            Workbook workbook = excelApplication.Workbooks.Open(excelPath);
            excelApplication.Visible = false;

            return workbook;
        }

        private List<Worksheet> GetWorkSheets(Workbook excelWorkbook)
        {
            List<Worksheet> workSheets = new List<Worksheet>();

            foreach (Worksheet worksheet in excelWorkbook.Worksheets)
            {
                if (!worksheet.Name.StartsWith("#"))
                {
                    workSheets.Add(worksheet);
                }
            }

            return workSheets;
        }

        private List<ListObject> GetListObjects(List<Worksheet> workSheets)
        {
            List<ListObject> listObjects = new List<ListObject>();

            foreach (Worksheet workSheet in workSheets)
            {
                foreach (ListObject listObject in workSheet.ListObjects)
                {
                    if (!listObject.Name.StartsWith("#"))
                    {
                        listObjects.Add(listObject);
                    }
                }
            }

            return listObjects;
        }

        private void ListObjectsToRawDataTable(List<ListObject> listObjects,
            ref System.Data.DataSet rawDataSet)
        {
            foreach (var listObj in listObjects)
            {
                System.Data.DataTable dataTable = new System.Data.DataTable(listObj.Name.Trim());

                // Add Columns
                int columnLength = listObj.HeaderRowRange.Columns.Count;
                for (int i = 1; i < columnLength + 1; ++i)
                {
                    string columnName = listObj.HeaderRowRange[1, i].Text.Trim();

                    dataTable.Columns.Add(new System.Data.DataColumn(columnName, typeof(string)));
                }

                // Add Rows(DataBody)
                int rowCount = listObj.DataBodyRange.Rows.Count;
                for (int rowOneBaseIndex = 1; rowOneBaseIndex < rowCount + 1; ++rowOneBaseIndex)
                {
                    string[] rowStrings = new string[columnLength];

                    for (int columnOneBaseIndex = 1; columnOneBaseIndex < columnLength + 1;
                        ++columnOneBaseIndex)
                    {
                        string value =
                            listObj.DataBodyRange[rowOneBaseIndex, columnOneBaseIndex].Text.Trim();

                        rowStrings[columnOneBaseIndex - 1] = value;
                    }

                    if (!Array.TrueForAll(rowStrings,
                        (x) => { return (x == null) || (x.Length == 0); }))
                    {
                        System.Data.DataRow dataRow = dataTable.NewRow();
                        dataRow.ItemArray = rowStrings;
                        dataTable.Rows.Add(dataRow);
                    }
                }

                rawDataSet.Tables.Add(dataTable);
            }
        }

        private void Release(object obj)
        {
            // Errors are ignored per Microsoft's suggestion for this type of function:
            // http://support.microsoft.com/default.aspx/kb/317109
            try
            {
                Marshal.ReleaseComObject(obj);
            }
            catch { }
        }

        private void SetFixedSystemTime()
        {
            if (!IsFixedDate)
                return;

            string appData = System.Environment.GetEnvironmentVariable("APPDATA");
            string excelPath = Path.Combine(appData, "Microsoft", "Excel");

            if (!Directory.Exists(excelPath))
                return;

            DateTime fixedDateTime = Directory.GetCreationTime(excelPath);
            fixedDateTime = fixedDateTime.AddDays(1);
            DateTime dateTime = DateTime.UtcNow;

            SYSTEMTIME systemTime = new SYSTEMTIME();

            systemTime.wYear = (ushort)fixedDateTime.Year;
            systemTime.wMonth = (ushort)fixedDateTime.Month;
            systemTime.wDay = (ushort)fixedDateTime.Day;
            systemTime.wDayOfWeek = (ushort)dateTime.DayOfWeek;
            systemTime.wHour = (ushort)dateTime.Hour;
            systemTime.wMinute = (ushort)dateTime.Minute;
            systemTime.wSecond = (ushort)dateTime.Second;
            systemTime.wMilliseconds = (ushort)dateTime.Millisecond;

            SetSystemTime(ref systemTime);
        }

        private void SetCurrentSystemTime()
        {
            if (!IsFixedDate)
                return;

            DateTime dateTime = DateTime.UtcNow;

            SYSTEMTIME systemTime = new SYSTEMTIME();

            systemTime.wYear = (ushort)CurrentYear;
            systemTime.wMonth = (ushort)CurrentMonth;
            systemTime.wDay = (ushort)CurrentDay;
            systemTime.wDayOfWeek = (ushort)dateTime.DayOfWeek;
            systemTime.wHour = (ushort)dateTime.Hour;
            systemTime.wMinute = (ushort)dateTime.Minute;
            systemTime.wSecond = (ushort)dateTime.Second;
            systemTime.wMilliseconds = (ushort)dateTime.Millisecond;

            SetSystemTime(ref systemTime);
        }

        public string ExcelPath { get; set; }
        public bool IsFixedDate { get; set; }
        public int CurrentYear { get; set; }
        public int CurrentMonth { get; set; }
        public int CurrentDay { get; set; }
    }
}
