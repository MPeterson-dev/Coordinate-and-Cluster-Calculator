using Microsoft.VisualBasic.ApplicationServices;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Coordinate_and_Cluster_Calculator.Models
{
    public class GridEntry
    {
        // The grid identifier, e.g., "N10-W4", "S5-2"
        public string GridCode { get; set; }

        // Quantity associated with this grid
        public double Quantity { get; set; }

        // Parsed X coordinate, used for positioning and distance calculations
        public int X { get; private set; }

        // Parsed Y coordinate, used for positioning and distance calculations
        public int Y { get; private set; }

        /*
         Quadrant Ranges:
            Northwest (N#-W#)       Northeast (#-#)
            
            -1, 1                   1, 1            
            
            Southwest (S#-W#)       Southeast (S#-#)
            
            -1, -1                  1, -1
         */
        public void ParsePosition()
        {
            var code = GridCode.ToUpper();
            bool south = code.StartsWith("S");
            code = code.TrimStart('N', 'S');

            var parts = code.Split('-');
            if (parts.Length != 2) throw new FormatException("Invalid GridCode format.");

            int row = int.Parse(parts[0]);
            int col = int.Parse(parts[1].TrimStart('W'));
            bool west = code.Contains("W");

            int xSign = west ? -1 : 1;
            int ySign = south ? -1 : 1;

            X = col * xSign;
            Y = row * ySign;
        }

        //Changes Parsed grids (-1, 1), (-1, -1) etc back to N#-W#, S#-W#, etc.
        public string ToGridCode()
        {
            int row = Math.Abs(Y);
            int col = Math.Abs(X);

            string ns = Y > 0 ? "N" : "S";
            string ew = X < 0 ? $"W{col}" : $"{col}";

            return $"{ns}{row}-{ew}";
        }

        /// <summary>
        /// Calculates the Manhattan distance to another grid.
        /// </summary>
        public double DistanceTo(GridEntry other)
        {
            return Math.Abs(this.X - other.X) + Math.Abs(this.Y - other.Y);
        }

        //Used in GetClustersandSolos
        public bool IsNeighbor(GridEntry other, int maxRange = 2)
        {
            int dx = Math.Abs(this.X - other.X);
            int dy = Math.Abs(this.Y - other.Y);
            return dx <= maxRange && dy <= maxRange;
        }        

        /*
         * Gets a list of the Grids and a range. CLusters both based on proximity using breadth-first search.
         * Each unvisited grid starts a new cluster.
         * Grids that are within the specified range of each other are grouped. Cluster = 2 or more. No neighbors = solo.
         */
        public static (List<List<GridEntry>> clusters, List<GridEntry> solos) GetClustersAndSolos(List<GridEntry> entries, int range)
        {
            var visited = new HashSet<GridEntry>();
            var clusters = new List<List<GridEntry>>();
            var solos = new List<GridEntry>();

            foreach (var entry in entries)
            {
                if (visited.Contains(entry))
                    continue;

                var cluster = new List<GridEntry>();
                var queue = new Queue<GridEntry>();
                queue.Enqueue(entry);
                visited.Add(entry);

                while (queue.Count > 0)
                {
                    var current = queue.Dequeue();
                    cluster.Add(current);

                    foreach (var neighbor in entries)
                    {
                        if (!visited.Contains(neighbor) && current.IsNeighbor(neighbor, range))
                        {
                            visited.Add(neighbor);
                            queue.Enqueue(neighbor);
                        }
                    }
                }

                if (cluster.Count > 1)
                {
                    clusters.Add(cluster);
                }
                else
                {
                    solos.Add(entry);
                }
            }

            return (clusters, solos);
        }

    }
}
