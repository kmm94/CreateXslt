namespace CreateXslt
{
    public enum Datatype
    {
        String,
        Number,
        Date
    }
    
    public class Column
    {
        public string sqlQueryHeadline { get; set; }
        public string name { get; set; } 
        public Datatype? datatype { get; set; }

        public Column(string sqlQueryHeadline, string name, Datatype? datatype = null)
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