using System;

namespace CreateXslt
{
    public enum ExcelFilters
    {
        String,
        Number,
        Date
    }
    public enum CrmReportInputType
    {
        Text,
        Dato,
        Decimal,
        Int,
        Long,
    }
    
    public class Column
    {
        public string _sqlQueryHeadline { get; private set; }
        public string userColumnTitle { get; set; }
        
        public ExcelFilters? datatype { get; set; }
        
        public CrmReportInputType CrmReportInputType { get; set; }

        public Column(string sqlQueryHeadline, ExcelFilters? datatype = null)
        {
            this._sqlQueryHeadline = sqlQueryHeadline;
            this.userColumnTitle = GuessTimateColumnTitle(sqlQueryHeadline);
            this.datatype = datatype;
        }

        private string GuessTimateColumnTitle(string sqlQueryHeadline)
        {
            var guesstimatedTitle = sqlQueryHeadline;

            if (guesstimatedTitle.IndexOf('_') == 2 && guesstimatedTitle.StartsWith("dm_", StringComparison.InvariantCultureIgnoreCase))
            {
                guesstimatedTitle = guesstimatedTitle.Substring(3);
            }

            guesstimatedTitle = guesstimatedTitle.Replace('_', ' ');
            guesstimatedTitle = guesstimatedTitle.Replace("oe", "ø");
            guesstimatedTitle = guesstimatedTitle.Replace("aa", "å");
            guesstimatedTitle = guesstimatedTitle.Replace("ae", "æ");
            
            guesstimatedTitle = char.ToUpper(guesstimatedTitle[0]) + guesstimatedTitle.Substring(1);
            return guesstimatedTitle;
        }

        public override string ToString()
        {
            return _sqlQueryHeadline;
        }
    }
}