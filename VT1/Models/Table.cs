namespace VT1.Models
{
    public class Table
    {
        public string name { get; set; }
        public string[] columns { get; set; }
        public string[,] values { get; set; }
        public int ? rowCount { get; set; }
        public int ? columnCount { get; set; }
    }
}
