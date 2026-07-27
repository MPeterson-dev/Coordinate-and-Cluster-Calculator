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
            setGraphPreferences();

            // Set up initial DataGridView columns
            dataGridView1.ColumnCount = 2;
            dataGridView1.Columns[0].HeaderText = "Grid Code";
            dataGridView1.Columns[1].HeaderText = "Quantity";
            dataGridView1.Columns[1].DefaultCellStyle.Format = "N0";

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
            else
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
            // Triggers plotting of grid entries from the view.
            if (allGrids.Count == 0)
            {
                MessageBox.Show("No grid data to plot. Please paste data first.");
                return;
            }

            var visibleEntries = GetVisibleEntriesFromGrid();
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

    }
}
