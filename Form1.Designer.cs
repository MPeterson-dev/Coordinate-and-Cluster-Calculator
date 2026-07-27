// Fully restored Form1.Designer.cs with regular WinForms buttons and all UI components
using System;
using System.Drawing;
using System.Windows.Forms;
using ScottPlot.WinForms;

namespace Coordinate_and_Cluster_Calculator
{
    partial class Form1
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
                components.Dispose();
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private void InitializeComponent()
        {
            label1 = new Label();
            comboFilterType = new ComboBox();
            textFilterValue = new TextBox();
            dataGridView1 = new DataGridView();
            buttonPaste = new Button();
            buttonClear = new Button();
            buttonFilter = new Button();
            buttonExport = new Button();
            numericNeighborRange = new NumericUpDown();
            label2 = new Label();
            label3 = new Label();
            label4 = new Label();
            label5 = new Label();
            label6 = new Label();
            label8 = new Label();
            label7 = new Label();
            formsPlot1 = new FormsPlot();
            buttonPlotGrids = new Button();
            panel1 = new Panel();
            ((System.ComponentModel.ISupportInitialize)dataGridView1).BeginInit();
            ((System.ComponentModel.ISupportInitialize)numericNeighborRange).BeginInit();
            panel1.SuspendLayout();
            SuspendLayout();
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Font = new Font("MS Reference Sans Serif", 15.75F, FontStyle.Bold);
            label1.Location = new Point(712, 199);
            label1.Name = "label1";
            label1.Size = new Size(118, 26);
            label1.TabIndex = 17;
            label1.Text = "Filter by:";
            // 
            // comboFilterType
            // 
            comboFilterType.DropDownStyle = ComboBoxStyle.DropDownList;
            comboFilterType.Font = new Font("Segoe UI", 12F);
            comboFilterType.Items.AddRange(new object[] { "Greater than", "Less than", "No filter" });
            comboFilterType.Location = new Point(825, 196);
            comboFilterType.Name = "comboFilterType";
            comboFilterType.Size = new Size(99, 29);
            comboFilterType.TabIndex = 16;
            // 
            // textFilterValue
            // 
            textFilterValue.Font = new Font("Segoe UI", 12F);
            textFilterValue.Location = new Point(930, 196);
            textFilterValue.Name = "textFilterValue";
            textFilterValue.PlaceholderText = "Amount";
            textFilterValue.Size = new Size(76, 29);
            textFilterValue.TabIndex = 15;
            textFilterValue.TextAlign = HorizontalAlignment.Center;
            // 
            // dataGridView1
            // 
            dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView1.Location = new Point(445, 143);
            dataGridView1.Name = "dataGridView1";
            dataGridView1.Size = new Size(250, 346);
            dataGridView1.TabIndex = 14;
            // 
            // buttonPaste
            // 
            buttonPaste.BackColor = SystemColors.MenuHighlight;
            buttonPaste.Font = new Font("Segoe UI", 12F);
            buttonPaste.ForeColor = SystemColors.ButtonHighlight;
            buttonPaste.Location = new Point(475, 98);
            buttonPaste.Name = "buttonPaste";
            buttonPaste.Size = new Size(193, 36);
            buttonPaste.TabIndex = 13;
            buttonPaste.Text = "Paste from Clipboard";
            buttonPaste.UseVisualStyleBackColor = false;
            buttonPaste.Click += buttonPaste_Click;
            // 
            // buttonClear
            // 
            buttonClear.BackColor = Color.Red;
            buttonClear.Font = new Font("Segoe UI", 12F);
            buttonClear.ForeColor = Color.White;
            buttonClear.Location = new Point(502, 496);
            buttonClear.Name = "buttonClear";
            buttonClear.Size = new Size(140, 36);
            buttonClear.TabIndex = 12;
            buttonClear.Text = "Clear Columns";
            buttonClear.UseVisualStyleBackColor = false;
            buttonClear.Click += buttonClear_Click;
            // 
            // buttonFilter
            // 
            buttonFilter.BackColor = SystemColors.MenuHighlight;
            buttonFilter.Font = new Font("Segoe UI", 12F);
            buttonFilter.ForeColor = SystemColors.ButtonHighlight;
            buttonFilter.Location = new Point(825, 263);
            buttonFilter.Name = "buttonFilter";
            buttonFilter.Size = new Size(68, 36);
            buttonFilter.TabIndex = 11;
            buttonFilter.Text = "Filter";
            buttonFilter.UseVisualStyleBackColor = false;
            buttonFilter.Click += buttonFilter_Click;
            // 
            // buttonExport
            // 
            buttonExport.BackColor = SystemColors.MenuHighlight;
            buttonExport.Font = new Font("Segoe UI", 12F);
            buttonExport.ForeColor = SystemColors.ButtonHighlight;
            buttonExport.Location = new Point(786, 332);
            buttonExport.Name = "buttonExport";
            buttonExport.Size = new Size(147, 36);
            buttonExport.TabIndex = 10;
            buttonExport.Text = "Export to Excel";
            buttonExport.UseVisualStyleBackColor = false;
            buttonExport.Click += buttonExport_Click;
            // 
            // numericNeighborRange
            // 
            numericNeighborRange.Font = new Font("Segoe UI", 12F);
            numericNeighborRange.Location = new Point(955, 141);
            numericNeighborRange.Minimum = new decimal(new int[] { 1, 0, 0, 0 });
            numericNeighborRange.Name = "numericNeighborRange";
            numericNeighborRange.Size = new Size(51, 29);
            numericNeighborRange.TabIndex = 9;
            numericNeighborRange.TextAlign = HorizontalAlignment.Center;
            numericNeighborRange.Value = new decimal(new int[] { 1, 0, 0, 0 });
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Font = new Font("MS Reference Sans Serif", 15.75F, FontStyle.Bold);
            label2.Location = new Point(712, 143);
            label2.Name = "label2";
            label2.Size = new Size(221, 26);
            label2.TabIndex = 8;
            label2.Text = "Set cluster range:";
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Font = new Font("Segoe UI", 12F);
            label3.Location = new Point(21, 50);
            label3.Name = "label3";
            label3.Size = new Size(324, 21);
            label3.TabIndex = 0;
            label3.Text = "1. Highlight two columns from an Excel sheet.";
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Font = new Font("Segoe UI", 12F);
            label4.Location = new Point(163, 15);
            label4.Name = "label4";
            label4.Size = new Size(47, 21);
            label4.TabIndex = 2;
            label4.Text = "Steps";
            label4.Click += label4_Click;
            // 
            // label5
            // 
            label5.AutoSize = true;
            label5.Font = new Font("Segoe UI", 12F);
            label5.Location = new Point(21, 71);
            label5.Name = "label5";
            label5.Size = new Size(262, 21);
            label5.TabIndex = 1;
            label5.Text = "2. Click Paste from Clipboard button.";
            // 
            // label6
            // 
            label6.AutoSize = true;
            label6.Font = new Font("Segoe UI", 12F);
            label6.Location = new Point(21, 92);
            label6.Name = "label6";
            label6.Size = new Size(189, 21);
            label6.TabIndex = 4;
            label6.Text = "3. Set ranges and/or filter.";
            // 
            // label8
            // 
            label8.AutoSize = true;
            label8.Font = new Font("Segoe UI", 12F);
            label8.Location = new Point(21, 113);
            label8.Name = "label8";
            label8.Size = new Size(365, 21);
            label8.TabIndex = 5;
            label8.Text = "4. Set Quantity filters (if needed). Click Filter button.";
            // 
            // label7
            // 
            label7.AutoSize = true;
            label7.Font = new Font("Segoe UI", 12F);
            label7.Location = new Point(21, 134);
            label7.Name = "label7";
            label7.Size = new Size(160, 21);
            label7.TabIndex = 3;
            label7.Text = "5. Click Export button.";
            // 
            // formsPlot1
            // 
            formsPlot1.DisplayScale = 1F;
            formsPlot1.Location = new Point(1044, 141);
            formsPlot1.Name = "formsPlot1";
            formsPlot1.Size = new Size(343, 346);
            formsPlot1.TabIndex = 1;
            // 
            // buttonPlotGrids
            // 
            buttonPlotGrids.BackColor = SystemColors.MenuHighlight;
            buttonPlotGrids.Font = new Font("Segoe UI", 12F);
            buttonPlotGrids.ForeColor = SystemColors.ButtonHighlight;
            buttonPlotGrids.Location = new Point(1200, 496);
            buttonPlotGrids.Name = "buttonPlotGrids";
            buttonPlotGrids.Size = new Size(64, 36);
            buttonPlotGrids.TabIndex = 0;
            buttonPlotGrids.Text = "Plot";
            buttonPlotGrids.UseVisualStyleBackColor = false;
            buttonPlotGrids.Click += buttonPlotGrids_Click;
            // 
            // panel1
            // 
            panel1.BackColor = SystemColors.ButtonHighlight;
            panel1.Controls.Add(label3);
            panel1.Controls.Add(label5);
            panel1.Controls.Add(label4);
            panel1.Controls.Add(label7);
            panel1.Controls.Add(label6);
            panel1.Controls.Add(label8);
            panel1.Location = new Point(23, 98);
            panel1.Name = "panel1";
            panel1.Size = new Size(397, 181);
            panel1.TabIndex = 18;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1421, 632);
            Controls.Add(buttonPlotGrids);
            Controls.Add(formsPlot1);
            Controls.Add(label2);
            Controls.Add(numericNeighborRange);
            Controls.Add(buttonExport);
            Controls.Add(buttonFilter);
            Controls.Add(buttonClear);
            Controls.Add(buttonPaste);
            Controls.Add(dataGridView1);
            Controls.Add(textFilterValue);
            Controls.Add(comboFilterType);
            Controls.Add(label1);
            Controls.Add(panel1);
            Name = "Form1";
            Text = "Coordinate and Cluster Calculator";
            ((System.ComponentModel.ISupportInitialize)dataGridView1).EndInit();
            ((System.ComponentModel.ISupportInitialize)numericNeighborRange).EndInit();
            panel1.ResumeLayout(false);
            panel1.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Label label1;
        private ComboBox comboFilterType;
        private TextBox textFilterValue;
        private DataGridView dataGridView1;
        private Button buttonPaste;
        private Button buttonClear;
        private Button buttonFilter;
        private Button buttonExport;
        private NumericUpDown numericNeighborRange;
        private Label label2;
        private Label label3;
        private Label label4;
        private Label label5;
        private Label label6;
        private Label label8;
        private Label label7;
        private FormsPlot formsPlot1;
        private Button buttonPlotGrids;
        private Panel panel1;
    }
}
