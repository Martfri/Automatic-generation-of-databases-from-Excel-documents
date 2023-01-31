using Newtonsoft.Json;

namespace VT1.Models
{
    public class Table
    {
        public string tableName { get; set; }

        //[JsonIgnore]
        public string[] columns { get; set; }

        //[JsonProperty("columns")]
        public string[] columnsWithDatatType { get; set; }

        public object?[,]? values { get; set; }
        public int ? rowCount { get; set; }
        public int? columnCount { get; set; }
        
    }
}
