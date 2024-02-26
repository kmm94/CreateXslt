using System.Globalization;

namespace CreateXslt
{
    public enum ExcelFilter
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
        public ExcelFilter? excelFilter { get; set; }
        
        public string sqlSelectTitle { get; private set; }
        
        public List<string> _rawData { get; }
        public CrmReportInputType CrmReportInputType { get; set; }

        public Column(string sqlQueryHeadline, List<string> rawData)
        {
            _sqlQueryHeadline = sqlQueryHeadline;
            userColumnTitle = GuesstimateColumnTitle(sqlQueryHeadline);
            excelFilter = GuesstimateDataType(rawData);
            CrmReportInputType = GuesstimateReportInput(rawData);
            sqlSelectTitle = GenerateSqlSelect(this);
            _rawData = rawData;
        }

        private string GenerateSqlSelect(Column column)
        {
            //TODO add convert to decimal(16,2) when double
           return column._sqlQueryHeadline.ToLower();
        }

        private CrmReportInputType GuesstimateReportInput(List<string> list)
        {
            //TODO implelemt
            return CrmReportInputType.Text;
        }

        private ExcelFilter? GuesstimateDataType(List<string> list)
        {
            KeyValuePair<ExcelFilter, double> filterProbability = GetHigestExcelFilterProbability(list);

            if (filterProbability.Value == 1)
            {
                return filterProbability.Key;
            }

            return ExcelFilter.String;
        }

        private KeyValuePair<ExcelFilter, double> GetHigestExcelFilterProbability(List<string> list)
        {
            Dictionary<ExcelFilter, double> filterProbability = new Dictionary<ExcelFilter, double>
            {
                { ExcelFilter.Date, 0},
                { ExcelFilter.Number, 0}
            };
            List<ExcelFilter> checkFilters = new List<ExcelFilter>(filterProbability.Keys);
            foreach (ExcelFilter excelFilter in checkFilters)
            {
                filterProbability[excelFilter] = GetProbability(excelFilter, list);
            }
            
            return GetHigestProbability(filterProbability);
        }

        private KeyValuePair<T, double> GetHigestProbability<T>(Dictionary<T, double> filterProbabilities)
        {
            KeyValuePair<T, double> highestProbability = filterProbabilities.First();
            
            foreach (KeyValuePair<T, double> keyValuePair in filterProbabilities)
            {
                if (highestProbability.Value < keyValuePair.Value)
                {
                    highestProbability = keyValuePair;
                }
            }
            return highestProbability;
        }

        private double GetProbability(ExcelFilter excelFilter, List<string> list)
        {
            switch(excelFilter)
            {
                case ExcelFilter.Date:
                    return GetProbabilityForDate(list);
                case ExcelFilter.Number:
                    return GetProbabilityForNumber(list);
                case ExcelFilter.String:
                    return 100;
                default:
                    throw new Exception($"datatype not implemented: {nameof(excelFilter)}");
            }
        }

        private double GetProbabilityForNumber(List<string> list)
        {
            int hits = 0;
            foreach (var data in list)
            {
                if (Decimal.TryParse(data, out _) || long.TryParse(data, out _) || double.TryParse(data, out _) || int.TryParse(data, out _))
                {
                    hits++;
                }
            }

            return hits == 0 ? 0 : hits / list.Count;
        }

        private double GetProbabilityForDate(List<string> list)
        {
            int hits = 0;
            
            string[] formats= {"M/d/yyyy h:mm:ss tt", "M/d/yyyy h:mm tt", 
                "MM/dd/yyyy hh:mm:ss", "M/d/yyyy h:mm:ss", 
                "M/d/yyyy hh:mm tt", "M/d/yyyy hh tt", 
                "M/d/yyyy h:mm", "M/d/yyyy h:mm", 
                "MM/dd/yyyy hh:mm", "M/dd/yyyy hh:mm",
                
                "dd/mm/yyyy hh:mm:ss", "mm-yyyy", "dd/mm/yyyy"};
            
            foreach (var data in list)
            {
                DateTime date = new DateTime();
                if (DateTime.TryParseExact(data, formats,
                        new CultureInfo("da-DK"),
                        DateTimeStyles.None, out date))
                {
                    hits++;
                }
            }

            return hits == 0 ? 0 : hits / list.Count;
        }

        private string GuesstimateColumnTitle(string sqlQueryHeadline)
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