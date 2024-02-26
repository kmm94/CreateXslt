namespace CreateXslt
{
    public class ExcelWorksheet
    {
        public List<Column> columns { get; set; }
        public string name { get; set; }
        
        public string specialFirstLine { get; set; }

        public ExcelWorksheet(List<Column> columns, string name = "Data")
        {
            this.columns = columns;
            this.name = name;
        }

        public ExcelWorksheet(List<Column> tableColumns)
        {
            this.columns = tableColumns;
        }
    }
}