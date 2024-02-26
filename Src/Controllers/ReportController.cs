﻿using DataAccess;
using NStack;
using Row = DataAccess.Row;

namespace CreateXslt
{
    public class ReportController
    {
        private ExcelWorksheet _excelWorksheet;
        private readonly XmlHelper _xmlHelper = new XmlHelper();
        public Column selectedColumn { get; set; }

        public XmlHelper XmlHelper
        {
            get { return _xmlHelper; }
        }

        public void LoadCsv(DataTable dataTable)
        {
            var tableColumns = BuildColumns(dataTable);
            _excelWorksheet = new ExcelWorksheet(tableColumns);
        }

        public ExcelWorksheet GetExcelWorksheet()
        {
            return _excelWorksheet;
        }

        public void SetColumnExcelFilter(int changedArgsSelectedItem)
        {
            selectedColumn.excelFilter = (ExcelFilter)Enum.Parse(typeof(ExcelFilter),
                GetExcelfilters()[changedArgsSelectedItem].ToString());
        }

        public int GetIndexOfExcelfilters(ExcelFilter? excelFilter)
        {
            return excelFilter == null ? 0 : GetExcelfilters().IndexOf(excelFilter.ToString());
        }

        public List<ustring> GetExcelfilters()
        {
            return new List<ustring>()
            {
                nameof(ExcelFilter.String),
                nameof(ExcelFilter.Date),
                nameof(ExcelFilter.Number)
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
                columns.Add(new Column(columnName, GetRawDateFromColumn(dataTable, columnName)));
            }

            return columns;
        }

        private List<string> GetRawDateFromColumn(DataTable dataTable, string columnName)
        {
            var colIndex = dataTable.GetColumnIndex(columnName);
            List<string> rawDate = new List<string>();
            foreach (Row row in dataTable.Rows)
            {
                var colData = row.Values[colIndex];
                if (string.IsNullOrEmpty(colData) is false)
                {
                    rawDate.Add(colData);
                }
            }

            return rawDate;
        }

        public void SetColumnCrmInputType(int changedArgsSelectedItem)
        {
            selectedColumn.CrmReportInputType = (CrmReportInputType)Enum.Parse(typeof(CrmReportInputType),
                GetExcelfilters()[changedArgsSelectedItem].ToString());
        }
    }
}