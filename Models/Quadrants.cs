
namespace Coordinate_and_Cluster_Calculator.Models
{
    public class Quadrants
    {
        public static Dictionary<string, List<GridEntry>> GroupByQuadrant(List<GridEntry> entries)
        {
            // Categorizes GridEntry items into NW, NE, SW, SE based on code prefix
            var quadrants = new Dictionary<string, List<GridEntry>>
            {
                { "NW", new List<GridEntry>() },
                { "NE", new List<GridEntry>() },
                { "SW", new List<GridEntry>() },
                { "SE", new List<GridEntry>() }
            };

            foreach (var entry in entries)
            {
                string code = entry.GridCode.ToUpper();

                if (code.StartsWith("N") && code.Contains("W"))
                    quadrants["NW"].Add(entry);
                else if (!code.StartsWith("N") && !code.StartsWith("S") && !code.Contains("W"))
                    quadrants["NE"].Add(entry);
                else if (code.StartsWith("S") && code.Contains("W"))
                    quadrants["SW"].Add(entry);
                else if (code.StartsWith("S") && !code.Contains("W"))
                    quadrants["SE"].Add(entry);
            }

            return quadrants;
        }

        public static (int excelRow, int excelCol) GetExcelPosition(GridEntry entry, int maxRow, int maxCol)
        {
            // Converts a GridEntry to Excel row/col index based on quadrant geometry
            var code = entry.GridCode.ToUpper().Replace("N", "").Replace("S", "").Replace("W", "");
            var parts = code.Split('-');
            int row = int.Parse(parts[0]);
            int col = int.Parse(parts[1]);

            string g = entry.GridCode.ToUpper();

            bool isNW = g.StartsWith("N") && g.Contains("W");
            bool isNE = !g.StartsWith("N") && !g.StartsWith("S") && !g.Contains("W");
            bool isSW = g.StartsWith("S") && g.Contains("W");
            bool isSE = g.StartsWith("S") && !g.Contains("W");

            int excelRow, excelCol;

            if (isNW)
            {
                excelRow = maxRow - row + 1;
                excelCol = maxCol - col + 1;
            }
            else if (isNE)
            {
                excelRow = maxRow - row + 1;
                excelCol = col;
            }
            else if (isSW)
            {
                excelRow = row;
                excelCol = maxCol - col + 1;
            }
            else
            {
                excelRow = row;
                excelCol = col;
            }

            return (excelRow + 1, excelCol + 1); // Excel uses 1-based indexing
        }
    }
}
