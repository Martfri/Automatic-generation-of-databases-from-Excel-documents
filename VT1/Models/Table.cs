using Newtonsoft.Json;

namespace VT1.Models
{
    public class Table
    {
        public string tableName { get; set; }
        public string[] columns { get; set; }
        public object?[,]? values { get; set; }
        public int ? rowCount { get; set; }
        public int? columnCount { get; set; }
        
    }
}
