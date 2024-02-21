﻿using System.Collections.Generic;

namespace CreateXslt
{
    public class ExcelWorksheet
    {
        public List<Column> columns { get; set; }
        public string name { get; set; }
        
        
        public ExcelWorksheet(List<Column> columns, string name)
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