using System.Collections.Generic;

namespace CreateXslt
{
    public class Report
    {
        public List<Column> columns { get; set; }
        public string name { get; set; }
        
        
        public Report(List<Column> columns, string name)
        {
            this.columns = columns;
            this.name = name;
        }
    }
}