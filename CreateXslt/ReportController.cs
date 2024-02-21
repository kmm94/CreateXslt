using System;
using System.Collections.Generic;
using DataAccess;

namespace CreateXslt
{
    public class ReportController
    {
        private DataTable csvReport;
        private ExcelWorksheet _excelWorksheet;

        public void LoadCsv(DataTable dataTable)
        {
            this.csvReport = dataTable;
            var tableColumns = GetColumns(dataTable);
            _excelWorksheet = new ExcelWorksheet(tableColumns);
        }

        public ExcelWorksheet GetExcelWorksheet()
        {
            return _excelWorksheet;
        }

        private List<Column> GetColumns(DataTable dataTable)
        {
            List<Column> columns = new List<Column>();

            Console.WriteLine("Found columns:");
            foreach (string columnName in dataTable.ColumnNames)
            {
                Console.WriteLine(columnName);
                columns.Add(new Column(columnName, columnName));
            }

            return columns;
        }
    }
    
    
}