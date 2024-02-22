namespace CreateXslt
{
    public enum ExcelFilters
    {
        String,
        Number,
        Date
    }
    
    public class Column
    {
        public string sqlQueryHeadline { get; set; }
        public string name { get; set; } 
        public ExcelFilters? datatype { get; set; }

        public Column(string sqlQueryHeadline, string name, ExcelFilters? datatype = null)
        {
            this.sqlQueryHeadline = sqlQueryHeadline;
            this.name = name;
            this.datatype = datatype;
        }

        public override string ToString()
        {
            return sqlQueryHeadline;
        }
    }
}