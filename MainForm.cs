using OfficeOpenXml.Packaging.Ionic.Zlib;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;



namespace ExcelImport
{
    public partial class MainForm : Form
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }

        public List<DataTable> dataTables = new List<DataTable>();
        public string currentTable = "";
        private Button button1;
        private Button button2;
        public List<FilterColumn> filterColumns;

        public MainForm()
        {
            InitializeComponent();

            string fileName = "data.xml";
            string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileName);
            if (System.IO.File.Exists(filePath))
            {
                using (FileStream stream = new FileStream(filePath, FileMode.Open))
                {
                    XmlSerializer serializer = new XmlSerializer(typeof(List<DataTable>));
                    dataTables = (List<DataTable>)serializer.Deserialize(stream);
                }

                // Ensure primary key column in all DataTables after loading from XML
                EnsurePrimaryKeyColumnInDataTables();
            }
            else
            {
                //Add the tables to the list if the save file doesn't exist
                dataTables.Add(new DataTable("CONVERSION_TRANSLATE"));
                dataTables.Add(new DataTable("CONVERSION_TRANSLATE_RELATED"));
                dataTables.Add(new DataTable("CONVERSION_TRANSLATE_COV"));
                dataTables.Add(new DataTable("CONVERSION_TRANSLATE_COV_SCCOL"));
                dataTables.Add(new DataTable("CONVERSION_TRANSLATE_TABLES"));

                // Add primary key column to each DataTable
                foreach (DataTable dataTable in dataTables)
                {
                    AddPrimaryKeyColumnToDataTable(dataTable);
                }
            }

            // set an initial value for currentTable
            currentTable = "CONVERSION_TRANSLATE";

            // call LoadData() at the end of the MainForm() constructor
            LoadData();
        }






        private void ConversionTranslateButton_Click(object sender, EventArgs e)
        {
            currentTable = "CONVERSION_TRANSLATE";
            LoadData();
        }

        private void ConversionTranslateRelatedButton_Click(object sender, EventArgs e)
        {
            currentTable = "CONVERSION_TRANSLATE_RELATED";
            LoadData();
        }

        private void ConversionTranslateCovButton_Click(object sender, EventArgs e)
        {
            currentTable = "CONVERSION_TRANSLATE_COV";
            LoadData();
        }

        private void ConversionTranslateCovSccolButton_Click(object sender, EventArgs e)
        {
            currentTable = "CONVERSION_TRANSLATE_COV_SCCOL";
            LoadData();
        }

        private void ConversionTranslateTablesButton_Click(object sender, EventArgs e)
        {
            currentTable = "CONVERSION_TRANSLATE_TABLES";
            LoadData();
        }

        private void AddPrimaryKeyColumnToDataTable(DataTable dataTable)
        {
            DataColumn pkColumn = dataTable.Columns.Add("ID", typeof(int));
            pkColumn.AutoIncrement = true;
            pkColumn.AutoIncrementSeed = 1;
            pkColumn.AutoIncrementStep = 1;
            pkColumn.ColumnMapping = MappingType.Hidden;
        }

        private void EnsurePrimaryKeyColumnInDataTables()
        {
            foreach (DataTable dataTable in dataTables)
            {
                // Check if the primary key column exists
                if (!dataTable.Columns.Contains("ID"))
                {
                    // If not, add the primary key column
                    AddPrimaryKeyColumnToDataTable(dataTable);
                }
            }
        }

        public System.Windows.Forms.DataGridView dataGridView1;

        public void LoadData()
        {
            // Clear the current DataGridView
            dataGridView1.DataSource = null;

            // Get the current DataTable
            DataTable dataTable = dataTables.Find(dt => dt.TableName == currentTable);

            // Set the DataGridView DataSource to the DataTable
            if (dataTable != null)
            {
                dataGridView1.DataSource = dataTable.DefaultView;
            }

            // Resize the columns
            dataGridView1.AutoResizeColumns();

            // Initialize filters
            LoadFilters();
        }

        public void LoadFilters()
        {
            // Remove any previously added filter text boxes
            if (filterColumns != null)
            {
                foreach (FilterColumn filterColumn in filterColumns)
                {
                    Controls.Remove(filterColumn.FilterTextBox);
                }
                filterColumns.Clear();
            }

            // Add filter text boxes for each column
            foreach (DataGridViewColumn column in dataGridView1.Columns)
            {
                // Create a new FilterColumn object
                FilterColumn filterColumn = new FilterColumn(column);

                // Set the MainForm property of the FilterColumn object
                filterColumn.MainForm = this;

                // Add the FilterColumn object to the list
               // filterColumns.Add(filterColumn);

                // Set the location and size of the filter text box
                filterColumn.FilterTextBox.Location = new Point(column.DataGridView.Location.X + column.DataGridView.RowHeadersWidth + column.DividerWidth + column.Width - filterColumn.Width, column.DataGridView.Location.Y);
                filterColumn.FilterTextBox.Width = filterColumn.Width;

                // Set the AutoCompleteCustomSource for the filter text box based on the data in the column
                filterColumn.PopulateAutoCompleteSource(dataGridView1);

                // Add the filterTextChanged event handler to the filter text box
                filterColumn.FilterTextBox.TextChanged += new EventHandler(filterTextBox_TextChanged);

                // Add the filterKeyDown event handler to the filter text box
                filterColumn.FilterTextBox.KeyDown += new KeyEventHandler(filterTextBox_KeyDown);

                // Add the filter text box to the form
                Controls.Add(filterColumn.FilterTextBox);
            }
        }



        public void ApplyFilters()
        {
            DataTable dataTable = dataTables.Find(dt => dt.TableName == currentTable);

            if (filterColumns != null && dataTable != null)
            {
                string filterExpression = "";

                foreach (FilterColumn filterColumn in filterColumns)
                {
                    if (!string.IsNullOrEmpty(filterColumn.FilterText))
                    {
                        if (filterExpression != "") filterExpression += " AND ";
                        filterExpression += string.Format("[{0}] LIKE '%{1}%'", filterColumn.ColumnName, filterColumn.FilterText);
                    }
                }

                dataTable.DefaultView.RowFilter = filterExpression;
            }
        }

        private void filterTextBox_TextChanged(object sender, EventArgs e)
        {
            // Ignore the event if there are no filter columns
            if (filterColumns.Count == 0) return;

            TextBox textBox = (TextBox)sender;
            FilterColumn filterColumn = filterColumns.Find(fc => fc.FilterTextBox == textBox);
            filterColumn.FilterText = textBox.Text;
            ApplyFilters();
        }


        private void filterTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            // Ignore the event if there are no filter columns
            if (filterColumns.Count == 0) return;

            if (e.KeyCode == Keys.Tab)
            {
                // Get the filter column for the current filter text box
                TextBox filterTextBox = (TextBox)sender;
                FilterColumn filterColumn = filterColumns.Find(fc => fc.FilterTextBox == filterTextBox);

                // Populate the auto complete source for the filter column
                filterColumn.PopulateAutoCompleteSource(dataGridView1);

                // Set the text of the filter text box to the first auto complete suggestion
                if (filterColumn.AutoCompleteCustomSource.Count > 0)
                {
                    filterTextBox.AutoCompleteCustomSource = filterColumn.AutoCompleteCustomSource;
                    filterTextBox.Text = filterColumn.AutoCompleteCustomSource[0];
                }

                // Select the remaining text in the filter text box
                filterTextBox.SelectAll();

                // Mark the key event as handled
                e.Handled = true;
            }
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            // Save the data to a local file
            string fileName = "data.xml";
            string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileName);

            using (FileStream stream = new FileStream(filePath, FileMode.Create))
            {
                XmlSerializer serializer = new XmlSerializer(typeof(List<DataTable>));
                serializer.Serialize(stream, dataTables);
            }
        }


        public class FilterColumn
        {
            public DataGridViewColumn Column { get; }
            public string Header { get; set; }
            public int Width { get; set; }
            public TextBox TextBox { get; set; }
            public List<string> AutoCompleteItems { get; set; }
            public int ColumnIndex { get; set; }
            public string ColumnName { get; set; }
            public TextBox FilterTextBox { get; set; }
            public string FilterText { get; set; } = "";
            public AutoCompleteStringCollection AutoCompleteCustomSource { get; set; }
            public MainForm MainForm { get; set; }
            public object FilterColumns { get; private set; }
            public List<FilterColumn> filterColumns = new List<FilterColumn>();


            public FilterColumn(DataGridViewColumn column)
            {
                Column = column;
                Header = column.HeaderText;
                Width = column.Width;
                TextBox = new TextBox();
                AutoCompleteItems = new List<string>();
                AutoCompleteCustomSource = new AutoCompleteStringCollection();
                FilterTextBox = new TextBox();
                ColumnIndex = column.Index;
                ColumnName = column.Name;
                FilterText = string.Empty;
                TextBox.TextChanged += new EventHandler(FilterColumnTextBox_TextChanged);
                TextBox.KeyDown += new KeyEventHandler(FilterColumnTextBox_KeyDown);
                AutoCompleteCustomSource.AddRange(AutoCompleteItems.ToArray());
            }

            private void FilterColumnTextBox_TextChanged(object sender, EventArgs e)
            {
                TextBox textBox = (TextBox)sender;
                FilterText = textBox.Text;
                string[] autoCompleteItems = AutoCompleteItems.ToArray();
                FilterTextBox.AutoCompleteCustomSource.Clear();
                FilterTextBox.AutoCompleteCustomSource.AddRange(autoCompleteItems);
                if (FilterTextBox.AutoCompleteCustomSource.Count > 0)
                {
                    FilterTextBox.Text = FilterTextBox.AutoCompleteCustomSource[0];
                    FilterTextBox.Select(FilterTextBox.Text.Length, 0);
                }
            }

            private void FilterColumnTextBox_KeyDown(object sender, KeyEventArgs e)
            {
                if (e.KeyCode == Keys.Tab)
                {
                    // Get the filter column for the current filter text box
                    TextBox filterTextBox = (TextBox)sender;
                    FilterColumn filterColumn = MainForm.filterColumns.FirstOrDefault(fc => fc.FilterTextBox == filterTextBox);


                    // Populate the auto complete source for the filter column
                    filterColumn.PopulateAutoCompleteSource(MainForm.dataGridView1);

                    // Set the text of the filter text box to the first auto complete suggestion
                    if (filterColumn.AutoCompleteCustomSource.Count > 0)
                    {
                        FilterTextBox.AutoCompleteCustomSource = filterColumn.AutoCompleteCustomSource;
                        FilterTextBox.Text = filterColumn.AutoCompleteCustomSource[0];
                    }

                    // Select the remaining text in the filter text box
                    FilterTextBox.SelectAll();

                    // Mark the key event as handled
                    e.Handled = true;
                }
            }

            public void PopulateAutoCompleteSource(DataGridView dataGridView)
            {
                // Get the distinct values for the column
                IEnumerable<string> distinctValues = dataGridView.Rows
                    .Cast<DataGridViewRow>()
                    .Where(r => !r.IsNewRow)
                    .Select(r => r.Cells[ColumnName].Value?.ToString())
                    .Where(s => !string.IsNullOrEmpty(s))
                    .Distinct();

                // Set the AutoCompleteCustomSource for the filter text box
                AutoCompleteItems.Clear();
                AutoCompleteItems.AddRange(distinctValues.ToArray());
                AutoCompleteCustomSource.Clear();
                AutoCompleteCustomSource.AddRange(AutoCompleteItems.ToArray());
            }
        }






        private void ImportDataButton_Click(object sender, EventArgs e)
        {
            ImportForm importForm = new ImportForm(dataTables, this); // Pass the entire dataTables list and MainForm reference
            importForm.ShowDialog();
            LoadData();
        }


        private void AddRowButton_Click(object sender, EventArgs e)
        {
            // Check if the current table has been set
            if (string.IsNullOrEmpty(currentTable))
            {
                MessageBox.Show("Please select a table first.");
                return;
            }

            try
            {
                // Get the current DataTable
                DataTable dataTable = dataTables.Find(dt => dt.TableName == currentTable);

                // Add a new row
                DataRow newRow = dataTable.NewRow();
                dataTable.Rows.Add(newRow);

                // Refresh the DataGridView
                dataGridView1.Refresh();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error adding row: " + ex.Message);
            }
        }

        private void LoadFiltersButton_Click(object sender, EventArgs e)
        {
            LoadFilters();
        }


        private void DeleteRowButton_Click(object sender, EventArgs e)
        {
            // Check if a row has been selected
            if (dataGridView1.SelectedRows.Count == 0)
            {
                MessageBox.Show("Please select a row to delete.");
                return;
            }

            try
            {
                // Get the current DataTable
                DataTable dataTable = dataTables.Find(dt => dt.TableName == currentTable);

                // Check if there are any rows in the DataTable
                if (dataTable.Rows.Count == 0)
                {
                    throw new Exception("No rows to delete.");
                }

                // Remove the selected row
                if (dataGridView1.SelectedRows.Count > 0)
                {
                    DataRowView rowView = dataGridView1.SelectedRows[0].DataBoundItem as DataRowView;
                    if (rowView != null)
                    {
                        DataRow row = rowView.Row;
                        dataTable.Rows.Remove(row);
                    }
                }

                // Refresh the DataGridView
                dataGridView1.Refresh();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error deleting row: " + ex.Message);
            }
        }





        private void MainForm_Load(object sender, EventArgs e)
        {
            // Set the initial currentTable to CONVERSION_TRANSLATE
            currentTable = "CONVERSION_TRANSLATE";

            // Load the data into the DataGridView
            LoadData();

            // Initialize the filterColumns list and create the filter text boxes
            filterColumns = new List<FilterColumn>();

            for (int i = 0; i < dataGridView1.Columns.Count; i++)
            {
                DataGridViewColumn column = dataGridView1.Columns[i];

                // Create a new filter text box for the column
                FilterColumn filterColumn = new FilterColumn(column);
                filterColumns.Add(filterColumn);

                // Add the filter text box to the form
                Controls.Add(filterColumn.TextBox);
            }

            // Add the filterTextChanged event handler to the filter text boxes
            foreach (FilterColumn filterColumn in filterColumns)
            {
                filterColumn.TextBox.TextChanged += new EventHandler(filterTextBox_TextChanged);
            }
        }



        public void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            // Get the current DataTable
            DataTable dataTable = dataTables.Find(dt => dt.TableName == currentTable);

            // Get the edited cell value
            object newValue = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;

            // Get the primary key value of the edited row
            object primaryKeyValue = dataTable.Rows[e.RowIndex]["ID"];

            // Find the row in the DataTable with the matching primary key value
            DataRow rowToUpdate = dataTable.Rows.Find(primaryKeyValue);

            // Update the value in the row
            rowToUpdate[e.ColumnIndex] = newValue;
        }

        private System.Windows.Forms.Button ConversionTranslateButton;
        private System.Windows.Forms.Button ConversionTranslateRelatedButton;
        private System.Windows.Forms.Button ConversionTranslateCovButton;
        private System.Windows.Forms.Button ConversionTranslateCovSccolButton;
        private System.Windows.Forms.Button ConversionTranslateTablesButton;
        private System.Windows.Forms.Button ImportButton;
        private System.Windows.Forms.Button LoadFiltersButton;

        private System.Windows.Forms.Button AddRowButton;
        private System.Windows.Forms.Button DeleteRowButton;

        public object FilterColumns { get; private set; }

        private void InitializeComponent()
        {
            this.ConversionTranslateButton = new System.Windows.Forms.Button();
            this.ConversionTranslateRelatedButton = new System.Windows.Forms.Button();
            this.ConversionTranslateCovButton = new System.Windows.Forms.Button();
            this.ConversionTranslateCovSccolButton = new System.Windows.Forms.Button();
            this.ConversionTranslateTablesButton = new System.Windows.Forms.Button();
            this.ImportButton = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.AddRowButton = new System.Windows.Forms.Button();
            this.DeleteRowButton = new System.Windows.Forms.Button();
            this.LoadFiltersButton = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // ConversionTranslateButton
            // 
            this.ConversionTranslateButton.Location = new System.Drawing.Point(12, 12);
            this.ConversionTranslateButton.Name = "ConversionTranslateButton";
            this.ConversionTranslateButton.Size = new System.Drawing.Size(127, 23);
            this.ConversionTranslateButton.TabIndex = 0;
            this.ConversionTranslateButton.Text = "CONVERSION_TRANSLATE";
            this.ConversionTranslateButton.UseVisualStyleBackColor = true;
            this.ConversionTranslateButton.Click += new System.EventHandler(this.ConversionTranslateButton_Click);
            // 
            // ConversionTranslateRelatedButton
            // 
            this.ConversionTranslateRelatedButton.Location = new System.Drawing.Point(145, 12);
            this.ConversionTranslateRelatedButton.Name = "ConversionTranslateRelatedButton";
            this.ConversionTranslateRelatedButton.Size = new System.Drawing.Size(189, 23);
            this.ConversionTranslateRelatedButton.TabIndex = 1;
            this.ConversionTranslateRelatedButton.Text = "CONVERSION_TRANSLATE_RELATED";
            this.ConversionTranslateRelatedButton.UseVisualStyleBackColor = true;
            this.ConversionTranslateRelatedButton.Click += new System.EventHandler(this.ConversionTranslateRelatedButton_Click);
            // 
            // ConversionTranslateCovButton
            // 
            this.ConversionTranslateCovButton.Location = new System.Drawing.Point(340, 12);
            this.ConversionTranslateCovButton.Name = "ConversionTranslateCovButton";
            this.ConversionTranslateCovButton.Size = new System.Drawing.Size(170, 23);
            this.ConversionTranslateCovButton.TabIndex = 2;
            this.ConversionTranslateCovButton.Text = "CONVERSION_TRANSLATE_COV";
            this.ConversionTranslateCovButton.UseVisualStyleBackColor = true;
            this.ConversionTranslateCovButton.Click += new System.EventHandler(this.ConversionTranslateCovButton_Click);
            // 
            // ConversionTranslateCovSccolButton
            // 
            this.ConversionTranslateCovSccolButton.Location = new System.Drawing.Point(516, 12);
            this.ConversionTranslateCovSccolButton.Name = "ConversionTranslateCovSccolButton";
            this.ConversionTranslateCovSccolButton.Size = new System.Drawing.Size(211, 23);
            this.ConversionTranslateCovSccolButton.TabIndex = 3;
            this.ConversionTranslateCovSccolButton.Text = "CONVERSION_TRANSLATE_COV_SCCOL";
            this.ConversionTranslateCovSccolButton.UseVisualStyleBackColor = true;
            this.ConversionTranslateCovSccolButton.Click += new System.EventHandler(this.ConversionTranslateCovSccolButton_Click);
            // 
            // ConversionTranslateTablesButton
            // 
            this.ConversionTranslateTablesButton.Location = new System.Drawing.Point(692, 12);
            this.ConversionTranslateTablesButton.Name = "ConversionTranslateTablesButton";
            this.ConversionTranslateTablesButton.Size = new System.Drawing.Size(175, 23);
            this.ConversionTranslateTablesButton.TabIndex = 4;
            this.ConversionTranslateTablesButton.Text = "CONVERSION_TRANSLATE_TABLES";
            this.ConversionTranslateTablesButton.UseVisualStyleBackColor = true;
            this.ConversionTranslateTablesButton.Click += new System.EventHandler(this.ConversionTranslateTablesButton_Click);
            // 
            // ImportButton
            // 
            this.ImportButton.Location = new System.Drawing.Point(12, 41);
            this.ImportButton.Name = "ImportButton";
            this.ImportButton.Size = new System.Drawing.Size(75, 23);
            this.ImportButton.TabIndex = 5;
            this.ImportButton.Text = "Import";
            this.ImportButton.UseVisualStyleBackColor = true;
            this.ImportButton.Click += new System.EventHandler(this.ImportDataButton_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(12, 70);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.Size = new System.Drawing.Size(855, 366);
            this.dataGridView1.TabIndex = 6;
            // 
            // AddRowButton
            // 
            this.AddRowButton.Location = new System.Drawing.Point(93, 41);
            this.AddRowButton.Name = "AddRowButton";
            this.AddRowButton.Size = new System.Drawing.Size(75, 23);
            this.AddRowButton.TabIndex = 7;
            this.AddRowButton.Text = "Add Row";
            this.AddRowButton.UseVisualStyleBackColor = true;
            this.AddRowButton.Click += new System.EventHandler(this.AddRowButton_Click);
            // 
            // DeleteRowButton
            // 
            this.DeleteRowButton.Location = new System.Drawing.Point(174, 41);
            this.DeleteRowButton.Name = "DeleteRowButton";
            this.DeleteRowButton.Size = new System.Drawing.Size(75, 23);
            this.DeleteRowButton.TabIndex = 8;
            this.DeleteRowButton.Text = "Delete Row";
            this.DeleteRowButton.UseVisualStyleBackColor = true;
            this.DeleteRowButton.Click += new System.EventHandler(this.DeleteRowButton_Click);
            // 
            // LoadFiltersButton
            // 
            this.LoadFiltersButton.Location = new System.Drawing.Point(256, 41);
            this.LoadFiltersButton.Name = "LoadFiltersButton";
            this.LoadFiltersButton.Size = new System.Drawing.Size(75, 23);
            this.LoadFiltersButton.TabIndex = 9;
            this.LoadFiltersButton.Text = "Load Filters";
            this.LoadFiltersButton.UseVisualStyleBackColor = true;
            this.LoadFiltersButton.Click += new System.EventHandler(this.LoadFiltersButton_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1111, 688);
            this.Controls.Add(this.DeleteRowButton);
            this.Controls.Add(this.AddRowButton);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.ImportButton);
            this.Controls.Add(this.ConversionTranslateTablesButton);
            this.Controls.Add(this.ConversionTranslateCovSccolButton);
            this.Controls.Add(this.ConversionTranslateCovButton);
            this.Controls.Add(this.ConversionTranslateRelatedButton);
            this.Controls.Add(this.ConversionTranslateButton);
            this.Name = "MainForm";
            this.Text = "MainForm";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }
    }
}
