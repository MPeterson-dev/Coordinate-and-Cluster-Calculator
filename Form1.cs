using ClosedXML.Excel;
using Coordinate_and_Cluster_Calculator.Models;
using System.ComponentModel;
using ScottPlot.Colormaps;
using System.Drawing;
using ScottPlot;
using Color = ScottPlot.Color;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Coordinate_and_Cluster_Calculator
{
    public partial class Form1 : Form
    {
        // Holds all loaded grid entries
        private List<GridEntry> allGrids = new();
        private NumericUpDown numericRange;

        public Form1()
        {
            InitializeComponent();
            Shown += displayWelcomeMessage;
            setGraphPreferences();

            // Set up initial DataGridView columns
            dataGridView1.ColumnCount = 2;
            dataGridView1.Columns[0].HeaderText = "Grid Code";
            dataGridView1.Columns[1].HeaderText = "Quantity";
            dataGridView1.Columns[1].DefaultCellStyle.Format = "N0";

        }

        private void displayWelcomeMessage(object sender, EventArgs e)
        {
            MessageBox.Show(
                "Welcome to the Coordinate and Cluster Calculator.\n\n" +
                "This application was created after I was promoted to lead a team using Advanced Methane Leak Detection (AMLD) technology and Discover software, which had not previously been used in our office.\n\n" +
                "Because there was no established strategy for surveying with the new technology, I developed this program to analyze grid coordinates and gas-line footage.\n\n" +
                "This program helped target, identify, and export mapped out clusters of high footage areas to an Excel sheet. \n\n" +
                "The application imports Excel data (Grid code and its paired footage), allows the user to set how close grids must be to be grouped into the same cluster, separates isolated grids, plots work areas, and exports color-coded Excel maps for assignment to the technicians.",
                "Coordinate and Cluster Calculator",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);

        }

        private void setGraphPreferences()
        {
            // Configures styling and crosshair for the ScottPlot graph.
            // Style the ScottPlot plot only once here
            var plt = formsPlot1.Plot;

            plt.Axes.Title.Label.Text = "Grid Plot";
            plt.Axes.Left.Label.Text = "Y Coordinate";
            plt.Axes.Bottom.Label.Text = "X Coordinate";

            plt.Axes.Title.Label.FontSize = 16;
            plt.Axes.Left.Label.FontSize = 12;
            plt.Axes.Bottom.Label.FontSize = 12;

            var cross = formsPlot1.Plot.Add.Crosshair(0, 0);

            cross.LineWidth = 2;
            cross.LineColor = ScottPlot.Colors.Magenta;

            formsPlot1.Refresh();
        }

        private List<GridEntry> LoadFilteredEntriesFromGrid()
        {
            // Extracts grid data from the DataGridView and parses positions into GridEntry objects.
            var entries = new List<GridEntry>();

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.Cells[0].Value != null && row.Cells[1].Value != null)
                {
                    string gridCode = row.Cells[0].Value.ToString();
                    if (double.TryParse(row.Cells[1].Value.ToString(), out double quantity))
                    {
                        var entry = new GridEntry
                        {
                            GridCode = gridCode,
                            Quantity = quantity
                        };

                        try
                        {
                            entry.ParsePosition();
                            entries.Add(entry);
                        }
                        catch (Exception ex)
                        {
                            // Show message if parsing fails
                            MessageBox.Show($"Failed to parse grid: {gridCode}\n{ex.Message}");
                        }
                    }
                }
            }

            return entries;
        }

        private void buttonPaste_Click(object sender, EventArgs e)
        {
            // Parses clipboard tab-separated text and populates the grid.
            var clipboardText = Clipboard.GetText();
            var lines = clipboardText.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);

            dataGridView1.Rows.Clear();
            allGrids.Clear();

            foreach (var line in lines)
            {
                var parts = line.Split('\t');
                if (parts.Length >= 2 && double.TryParse(parts[1], out double quantity))
                {
                    var entry = new GridEntry
                    {
                        GridCode = parts[0].Trim(),
                        Quantity = quantity
                    };
                    entry.ParsePosition();
                    allGrids.Add(entry);
                    dataGridView1.Rows.Add(entry.GridCode, Math.Round(entry.Quantity, 0));

                }
            }
        }

        private void buttonClear_Click(object sender, EventArgs e)
        {
            // Clears all entries from the internal list and the grid view.
            allGrids.Clear();
            dataGridView1.Rows.Clear();
        }

        private List<GridEntry> GetFilteredData()
        {
            // Returns all entries currently visible in the DataGridView.
            return GetVisibleEntriesFromGrid();
        }

        private void CreateClusterAndSoloSheets(XLWorkbook workbook, List<GridEntry> filteredEntries, int neighborRange, out List<List<GridEntry>> clusters, out List<GridEntry> solos)
        {
            // Generates Excel worksheets for clusters and solos from filtered grid data.
            (clusters, solos) = GridEntry.GetClustersAndSolos(filteredEntries, neighborRange);

            var wsClusters = workbook.Worksheets.Add("Clusters");
            int colIndex = 1;
            foreach (var cluster in clusters)
            {
                wsClusters.Cell(1, colIndex).Value = "Grid";
                wsClusters.Cell(1, colIndex + 1).Value = "Quantity";
                int row = 2;
                foreach (var g in cluster)
                {
                    wsClusters.Cell(row, colIndex).Value = g.GridCode;
                    wsClusters.Cell(row, colIndex + 1).Value = g.Quantity;
                    row++;
                }
                colIndex += 3;
            }

            var wsSolos = workbook.Worksheets.Add("Solos");
            wsSolos.Cell(1, 1).Value = "Grid";
            wsSolos.Cell(1, 2).Value = "Quantity";
            int soloRow = 2;
            foreach (var g in solos)
            {
                wsSolos.Cell(soloRow, 1).Value = g.GridCode;
                wsSolos.Cell(soloRow, 2).Value = g.Quantity;
                soloRow++;
            }
        }

        private Dictionary<string, List<GridEntry>> CreateQuadrantTabs(XLWorkbook workbook, List<List<GridEntry>> clusters)
        {
            // Creates Excel sheets for each quadrant, arranged visually by coordinates.
            var clusterOnly = clusters.SelectMany(c => c).ToList();
            var quadrantGroups = Quadrants.GroupByQuadrant(clusterOnly);


            foreach (var kvp in quadrantGroups)
            {
                var sheet = workbook.Worksheets.Add($"Map_{kvp.Key}");
                var entries = kvp.Value;

                if (entries.Any())
                {
                    int maxRow = entries.Max(e => int.Parse(e.GridCode.Split('-')[0].TrimStart('N', 'S')));
                    int maxCol = entries.Max(e => int.Parse(e.GridCode.Split('-')[1].TrimStart('W')));


                    foreach (var entry in entries)
                    {
                        var (r, c) = Quadrants.GetExcelPosition(entry, maxRow, maxCol);
                        var cell = sheet.Cell(r, c);
                        cell.Value = $"{entry.GridCode}{Environment.NewLine}{entry.Quantity}";
                        cell.Style.Alignment.WrapText = true;
                        cell.Style.Fill.BackgroundColor = GetColorForQuantity(entry.Quantity);
                        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    }
                }

                sheet.Columns().Width = 10;
                sheet.Rows().Height = 30;

            }

            return quadrantGroups;
        }

        private void CreateCombinedMapSheet(XLWorkbook workbook, Dictionary<string, List<GridEntry>> quadrantGroups)
        {
            // Builds a unified map tab with all grid entries correctly placed.
            var sheet = workbook.Worksheets.Add("Map_All");

            // Step 1: Flatten all quadrant entries into one list
            var allEntries = quadrantGroups.SelectMany(kvp => kvp.Value).ToList();

            // Step 2: Determine X and Y coordinates
            int minX = allEntries.Min(e => e.X);
            int maxX = allEntries.Max(e => e.X);
            int minY = allEntries.Min(e => e.Y);
            int maxY = allEntries.Max(e => e.Y);

            // Step 3: Normalize and write each entry based on true geometry
            foreach (var entry in allEntries)
            {
                int excelCol = entry.X - minX + 1;           // X maps to column (left = minX)
                int excelRow = maxY - entry.Y + 1;           // Y maps to row (top = maxY)

                var cell = sheet.Cell(excelRow, excelCol);
                cell.Value = $"{entry.GridCode}{Environment.NewLine}{entry.Quantity}";
                cell.Style.Alignment.WrapText = true;
                cell.Style.Fill.BackgroundColor = GetColorForQuantity(entry.Quantity);
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            }

            sheet.Columns().Width = 10;
            sheet.Rows().Height = 30;

        }

        private XLColor GetColorForQuantity(double quantity)
        {
            // Returns a color based on quantity range using pre-defined thresholds.
            if (quantity <= 5000)
                return XLColor.Silver;
            else if (quantity <= 10000)
                return XLColor.LightYellow;
            else if (quantity <= 15000)
                return XLColor.LightGreen;
            else if (quantity <= 20000)
                return XLColor.LightBlue;
            else if (quantity <= 25000)
                return XLColor.Peach;
            else if (quantity <= 30000)
                return XLColor.Lavender;
            else if (quantity <= 35000)
                return XLColor.MistyRose;
            else // <35001
                return XLColor.BrightTurquoise;
        }

        public void buttonExport_Click(object sender, EventArgs e)
        {
            // Exports filtered data and visual map to an Excel file.
            int neighborRange = (int)numericNeighborRange.Value;
            double.TryParse(textFilterValue.Text, out double quantityFilter);

            using var sfd = new SaveFileDialog
            {
                Filter = "Excel Workbook (*.xlsx)|*.xlsx",
                FileName = $"Clusters_and_Solos_quantity_filter-{(comboFilterType.SelectedItem?.ToString()?.ToLower().Replace(" ", "") ?? "none")}-{quantityFilter}_Neighbor_range-{neighborRange}.xlsx"
            };

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                var filePath = sfd.FileName;
                var filteredEntries = GetFilteredData();

                var workbook = new XLWorkbook();

                CreateClusterAndSoloSheets(workbook, filteredEntries, neighborRange, out var clusters, out var solos);
                var quadrantGroups = CreateQuadrantTabs(workbook, clusters);
                CreateCombinedMapSheet(workbook, quadrantGroups);

                workbook.SaveAs(filePath);
                MessageBox.Show("Export complete!", "Success");
            }
        }

        private void buttonFilter_Click(object sender, EventArgs e)
        {
            // Filters grid view rows based on user input and filter type.
            string filterType = comboFilterType.SelectedItem?.ToString()?.ToLower() ?? "greater than";

            if (!double.TryParse(textFilterValue.Text, out double filterValue))
            {
                MessageBox.Show("Please enter a valid number.");
                return;
            }

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.IsNewRow) continue;

                if (row.Cells[1].Value != null && double.TryParse(row.Cells[1].Value.ToString(), out double qty))
                {
                    bool keep = filterType.Contains("no filter") ||
                                (filterType.Contains("greater") && qty >= filterValue) ||
                                (filterType.Contains("less") && qty <= filterValue);

                    row.Visible = keep;
                }
            }
        }

        private void PlotGrids(List<GridEntry> entries)
        {
            // Draws grid points on the plot with colors scaled to quantity.
            var plt = formsPlot1.Plot;

            plt.Clear();

            double max = entries.Max(e => e.Quantity);
            double min = entries.Min(e => e.Quantity);
            var cmap = new Turbo();

            foreach (var entry in entries)
            {
                double norm = (entry.Quantity - min) / (max - min + 1e-6);
                var color = cmap.GetColor(norm);

                var scatter = plt.Add.ScatterPoints(
                    xs: new[] { (double)entry.X },
                    ys: new[] { (double)entry.Y },
                    color: color
                );
                scatter.MarkerSize = 10;

            }

            plt.Axes.AutoScale();
            formsPlot1.Refresh();
        }

        private void label4_Click(object sender, EventArgs e)
        {
            string imagePath = Path.Combine(Application.StartupPath, "Pictures", "Jack.jpg");

            if (File.Exists(imagePath))
            {
                Form popup = new Form
                {
                    Text = "Here's your puppy!",
                    Size = new Size(400, 400),
                    StartPosition = FormStartPosition.CenterParent
                };

                PictureBox picture = new PictureBox
                {
                    Image = System.Drawing.Image.FromFile(imagePath),
                    SizeMode = PictureBoxSizeMode.Zoom,
                    Dock = DockStyle.Fill
                };

                popup.Controls.Add(picture);
                popup.ShowDialog();
            }
            else
            {
                MessageBox.Show("Couldn?t find 'Jack.jpg' in the Models folder.", "Image Missing", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void buttonPlotGrids_Click(object sender, EventArgs e)
        {
            MessageBox.Show(

                "ScottPlot was initially added to provide a visual representation " +
                "of grid locations and footage. While useful for quickly viewing " +
                "spatial relationships, the plot was less practical for creating " +
                "technician work assignments. Excel export was later added to provide " +
                "structured, color-coded maps that could be reviewed, organized, " +
                "and distributed more effectively.",
                "Plot Information",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);

            // Retrieves only the grid entries currently visible in the DataGridView.
            var visibleEntries = GetVisibleEntriesFromGrid();

            if (visibleEntries.Count == 0)
            {
                MessageBox.Show(
                    "No grid data is available to plot. Please paste or add sample data first.",
                    "Plot Grids",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                return;
            }

            PlotGrids(visibleEntries);
        }

        private List<GridEntry> GetVisibleEntriesFromGrid()
        {
            // Returns only visible grid entries parsed from the DataGridView.
            var visibleEntries = new List<GridEntry>();

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.Visible && row.Cells[0].Value != null && row.Cells[1].Value != null)
                {
                    string gridCode = row.Cells[0].Value.ToString();
                    if (double.TryParse(row.Cells[1].Value.ToString(), out double quantity))
                    {
                        var entry = new GridEntry
                        {
                            GridCode = gridCode,
                            Quantity = quantity
                        };

                        try
                        {
                            entry.ParsePosition();
                            visibleEntries.Add(entry);
                        }
                        catch
                        {
                            // Skip rows that fail parsing
                        }
                    }
                }
            }

            return visibleEntries;
        }

        private void buttonAddSampleData_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            allGrids.Clear();

            // tuple instead of dataGridView1.Rows.Add(...)
            var sampleData = new (string GridCode, double Quantity)[]
            {
                    // Northwest: large, high-footage cluster
                    ("N20-W20", 42000),
                    ("N20-W21", 38000),
                    ("N20-W22", 34000),
                    ("N21-W20", 29000),
                    ("N21-W21", 26000),
                    ("N21-W22", 23000),
                    ("N22-W21", 19000),
                    ("N22-W22", 14000),
            
                    // Northwest: nearby smaller cluster
                    // A larger neighbor range may connect this to the first cluster.
                    ("N24-W23", 11000),
                    ("N24-W24", 16000),
                    ("N25-W24", 21000),
            
                    // Northeast: horizontal cluster
                    ("N12-15", 7000),
                    ("N12-16", 12000),
                    ("N12-17", 18000),
                    ("N12-18", 24000),
                    ("N12-19", 31000),
            
                    // Northeast: isolated grids
                    ("N5-30", 4500),
                    ("N8-35", 36000),
            
                    // Southwest: irregular cluster
                    ("S10-W12", 9000),
                    ("S10-W13", 15000),
                    ("S11-W12", 20000),
                    ("S11-W13", 25000),
                    ("S11-W14", 30000),
                    ("S12-W14", 35000),
            
                    // Southwest: solo grid
                    ("S20-W30", 6000),
            
                    // Southeast: vertical cluster
                    ("S5-10", 5000),
                    ("S6-10", 10000),
                    ("S7-10", 15000),
                    ("S8-10", 20000),
                    ("S9-10", 25000),
            
                    // Southeast: separate high-footage cluster
                    ("S15-20", 28000),
                    ("S15-21", 33000),
                    ("S16-20", 39000),
                    ("S16-21", 46000),
            
                    // Southeast: isolated low-footage grid
                    ("S25-35", 2500)
            };

            //HashSet to avoid duplicate entries
            HashSet<string> usedGridCodes = new HashSet<string>();

            //Add Sample data to grid
            foreach (var entry in sampleData)
            {
                dataGridView1.Rows.Add(entry.GridCode, entry.Quantity);
                //Add sample data to HashSet 
                usedGridCodes.Add(entry.GridCode);
            }

            //Create random data for more samples
            Random random = new Random();

            //Add 2000 additional randomized grids
            while (usedGridCodes.Count < sampleData.Length + 2000){
                int x = random.Next(-100, 101);
                int y = random.Next(-100, 101);

                if (x == 0 || y == 0)
                    continue;

                string gridCode = GridEntry.ToGridCode(x, y);

                if (!usedGridCodes.Add(gridCode))
                    continue;

                double quantity = random.Next(1000, 50001);

                dataGridView1.Rows.Add(gridCode, quantity);
            }

            MessageBox.Show(
                "Sample data has been added.\n\n" +
                "Try filtering the data with different neighbor-range values to see " +
                "how nearby groups merge or remain separate.\n\n" +
                "The sample includes all four quadrants, multiple cluster shapes, " +
                "isolated grids, and footage values across every export color range.",
                "Sample Data Added",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);

        }
    }
}
