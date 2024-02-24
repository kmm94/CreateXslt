using System;
using System.Collections.Generic;
using DataAccess;
using NStack;

namespace CreateXslt
{
    public class ReportController
    {
        private DataTable csvReport;
        private ExcelWorksheet _excelWorksheet;
        public Column selectedColumn { get; set; }

        public void LoadCsv(DataTable dataTable)
        {
            this.csvReport = dataTable;
            var tableColumns = BuildColumns(dataTable);
            _excelWorksheet = new ExcelWorksheet(tableColumns);
        }

        public ExcelWorksheet GetExcelWorksheet()
        {
            return _excelWorksheet;
        }

        public void SetColumnExcelFilter(int changedArgsSelectedItem)
        {
            selectedColumn.datatype = (ExcelFilters) Enum.Parse(typeof(ExcelFilters),GetExcelfilters()[changedArgsSelectedItem].ToString());
        }
        
        public List<ustring> GetExcelfilters()
        {
            return new List<ustring>()
            {
                nameof(ExcelFilters.String),
                nameof(ExcelFilters.Date),
                nameof(ExcelFilters.Number)
            };
        }
        public List<ustring> GetCrmReportInputType()
        {
            return new List<ustring>()
            {
                nameof(CrmReportInputType.Text),
                nameof(CrmReportInputType.Dato),
                nameof(CrmReportInputType.Decimal),
                nameof(CrmReportInputType.Int),
                nameof(CrmReportInputType.Long)
            };
        }

        private List<Column> BuildColumns(DataTable dataTable)
        {
            List<Column> columns = new List<Column>();
            
            foreach (string columnName in dataTable.ColumnNames)
            {
                columns.Add(new Column(columnName));
            }

            return columns;
        }

        public void SetColumnCrmInputType(int changedArgsSelectedItem)
        {
            selectedColumn.CrmReportInputType = (CrmReportInputType) Enum.Parse(typeof(CrmReportInputType),GetExcelfilters()[changedArgsSelectedItem].ToString());
        }
    }
    
    
}