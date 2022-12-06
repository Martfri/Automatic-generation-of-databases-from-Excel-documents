using OfficeOpenXml;
using VT1.Models;
//ws.Cells[i - 1, j, i - 1, j].Value != null && ws.Cells[i - 1, j, i - 1, j].Value.GetType() == typeof(string) && ws.Cells[i, j, i, j].Value.GetType() == typeof(string) && !HaveSimilarFormat(i, j, i - 1, j) && ws.Cells[i, j + 1, i, j + 1].Value != null && HaveSimilarFormat(i, j, i, j + 1

namespace VT1.Services
{
    class ExcelService
    {
        private ExcelPackage Excelpackage;
        private ExcelWorksheet ws;
        private List<TableName> tables = new List<TableName>();
        private List<Column> columns = new List<Column>();
        private List<Entry> dataEntries = new List<Entry>();
        //string filePath = ".\\wwwroot\\Uploads\\Test.xlsx";


        public ExcelService(ExcelPackage package)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            Excelpackage = package;
            ws = Excelpackage.Workbook.Worksheets.First();
        }


        public List<TableName> TableDetection()
        {
            int colCount = ws.Dimension.End.Column;  //get Column Count
            int rowCount = ws.Dimension.End.Row;     //get row count
            
            for (int row = rowCount; row >= 1; row--)
            {
                for (int col = colCount; col >= 1; col--)
                {
                    TableRecognition(row, col);
                }
            }
            return tables;
        }

        public List<TableName> TableRecognition(int i, int j)
        {          

           if (ws.Cells[i, j].Value == null) return null;

           else if (IsDataCell(i, j))
           {

                var data = new Entry();
                data.i = i;
                data.j = j;
                data.value = ws.Cells[i, j].Value.ToString();
                dataEntries.Add(data);
           }

           else if (IsHeaderCell(i, j))
           {
                var column = new Column();
                column.count = j;
                column.Name = ws.Cells[i, j].Value.ToString();                
                column.entries = dataEntries.Where(p => p.j == j).ToList();
                columns.Add(column);
           }

           else if (IsTitleCell(i, j))
           {
                var title = new TableName();
                title.name = ws.Cells[i, j].Value.ToString();
                title.columns = columns;
                tables.Add(title);
           }

           else { return null; }

           return tables;


        }
        
        public bool IsTitleCell(int i, int j)
        {
            if (ws.Cells[i, j].Value.GetType() == typeof(string) && columns.Count != 0 && ws.Cells[i + 1, j].Value == null)
            {
                return true;
            }

            else if(tables.Count != 0 && (IsTitleCell(i+1, j)|| IsTitleCell(i + 1, j-1) || IsTitleCell(i + 1, j +1))) 
            {
                return true;
            }

            else if (columns.Count != 0 && ws.Cells[i, j+1].Value == null && ws.Cells[i, j - 1].Value == null && columns.First().count == j)
            {
                return true;
            }

            else
            {
                return false;
            }
        }

        public bool IsHeaderCell(int i, int j)
        {
            if (ws.Cells[i, j].Value == null) return false;

            else if (columns.Count != 0 && ws.Cells[i, j].Value.GetType() == typeof(string) && ws.Cells[i+1,j].Value != null)
            {
                return true;
            }
                //check also for color
            else if (HasBorders(i, j) && (HasBorders(i, j + 1) || HasBorders(i, j - 1)) && ws.Cells[i + 1, j].Value != null)
            {
                return true;
            }

            else if (HaveSimilarFormat(i, j, i, j - 1) && HaveSimilarFormat(i, j, i, j + 1) && ws.Cells[i + 1, j].Value != null)
            {
                 return true;
            }

             //nochmals checken
            else if (ws.Cells[i + 1,j].Value != null && ws.Cells[i + 1,j].Value.GetType() != typeof(string) && ws.Cells[i,j].Value.GetType() == typeof(string) && !HaveSimilarFormat(i, j, i + 1, j) && ws.Cells[i, j - 1].Value != null && HaveSimilarFormat(i, j, i, j - 1))
            {
                return true;
            }

            else
            {
                return false;
            }
        }

        public bool IsDataCell(int i, int j)
        {
            if (HaveSimilarFormat(i, j, i - 1, j) && HaveSimilarFormat(i, j, i + 1, j))
            {
                return true;
            }

            else if (HaveSimilarFormat(i, j, i - 1, j) && ws.Cells[i + 1, j].Value == null)            
            {
                return true;
            }

            else if (HaveSimilarFormat(i, j, i +1, j) && IsHeaderCell(i -1, j) == true)
            {
                return true;
            }

            else
            {
                return false;
            }
        }

        public bool HaveSimilarFormat(int i, int j, int ii, int jj)
        { 
            if ((ii == 0) || (jj == 0) || (ws.Cells[ii, jj].Value == null)) return false;
            

            else if (ws.Cells[i, j].Style.Font.Size == ws.Cells[ii, jj].Style.Font.Size && ws.Cells[i, j].Style.Font.Color.Tint == ws.Cells[ii, jj].Style.Font.Color.Tint)
            {
                return true;
            }

            else
            {
                return false;
            }
        }

        public bool HasBorders(int i, int j)
        {
            if ((i == 0) || (j == 0) || (ws.Cells[i, j].Value == null)) return false;

            else if (ws.Cells[i, j].Style.Border.Top != null || ws.Cells[i, j].Style.Border.Bottom != null || ws.Cells[i, j].Style.Border.Left != null || ws.Cells[i, j].Style.Border.Right != null)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public Table MapToTable(TableName tableName)
        {
            int colCount = tableName.columns.Count;
            int rowCount = tableName.columns.Select(x => x.entries.Count).Max();

            Table table = new Table();
            table.name = tableName.name;
            table.columns = new string[colCount];
            table.values = new string[rowCount, colCount];
            table.columnCount = colCount;
            table.rowCount = rowCount;

            var columns = tableName.columns.OrderBy(x => x.count).ToList();
            for (int colIdx = 0; colIdx < columns.Count; colIdx++)
            {
                var col = columns[colIdx];
                table.columns[colIdx] = col.Name;

                for (int rowIdx = 0; rowIdx < col.entries.Count; rowIdx++)
                {
                    var entry = col.entries[rowIdx];

                    // falls Null-Zellen im Excel erlaubt sind, muss dies hier noch implementiert werden!
                    table.values[rowIdx, colIdx] = entry.value;
                }
            }

            return table;
        }
    }
}



