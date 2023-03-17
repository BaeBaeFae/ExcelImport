using System;
using System.IO;
using System.Data;
using System.Windows.Forms;
using OfficeOpenXml;
using System.Data.OleDb;
using System.Collections.Generic;

namespace ExcelImport
{
    public partial class ImportForm : Form
    {
        public List<DataTable> dataTables = new List<DataTable>();

        public ImportForm()
        {
            InitializeComponent();
           // dataGridView1.DataSource = dataTables[0];

        }

        private void ImportButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            openFileDialog.Title = "Select an Excel file";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog.FileName;
                string[] sheetNames = GetExcelSheetNames(filePath);

                foreach (string sheetName in sheetNames)
                {
                    DataTable dt = ReadExcelFile(filePath, sheetName);
                    dataTables.Add(dt);
                }

                dataGridView1.DataSource = dataTables[0];
            }
        }
        private DataTable ReadExcelFile(string filePath, string sheetName)
        {
            DataTable dt = new DataTable(sheetName);

            using (OleDbConnection connection = new OleDbConnection())
            {
                connection.ConnectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1\"", filePath);
                using (OleDbCommand command = new OleDbCommand())
                {
                    command.CommandText = "SELECT * FROM [" + sheetName + "$]";
                    command.Connection = connection;
                    using (OleDbDataAdapter adapter = new OleDbDataAdapter())
                    {
                        adapter.SelectCommand = command;
                        adapter.Fill(dt);
                    }
                }
            }

            // Add hidden primary key column
            DataColumn pkColumn = dt.Columns.Add("ID", typeof(int));
            pkColumn.AutoIncrement = true;
            pkColumn.AutoIncrementSeed = 1;
            pkColumn.AutoIncrementStep = 1;
            pkColumn.ColumnMapping = MappingType.Hidden;

            return dt;
        }

        private string[] GetExcelSheetNames(string filePath)
        {
            DataTable dt = new DataTable();

            using (OleDbConnection connection = new OleDbConnection())
            {
                connection.ConnectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1\"", filePath);
                connection.Open();
                dt = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                connection.Close();
            }

            string[] sheetNames = new string[dt.Rows.Count];
            int i = 0;

            foreach (DataRow row in dt.Rows)
            {
                sheetNames[i] = row["TABLE_NAME"].ToString().Replace("$", "");
                i++;
            }

            return sheetNames;
        }





        private void SaveButton_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "XML File|*.xml";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                DataSet ds = new DataSet();
                foreach (DataTable dt in dataTables)
                {
                    ds.Tables.Add(dt);
                }
                ds.WriteXml(sfd.FileName);
                MessageBox.Show("Data saved successfully.");
            }
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            // Close the ImportForm
            Close();
        }

        private void InitializeComponent()
        {
            this.ImportButton = new System.Windows.Forms.Button();
            this.SaveButton = new System.Windows.Forms.Button();
            this.CancelButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // ImportButton
            // 
            this.ImportButton.Location = new System.Drawing.Point(12, 12);
            this.ImportButton.Name = "ImportButton";
            this.ImportButton.Size = new System.Drawing.Size(75, 23);
            this.ImportButton.TabIndex = 0;
            this.ImportButton.Text = "Import";
            this.ImportButton.UseVisualStyleBackColor = true;
            this.ImportButton.Click += new System.EventHandler(this.ImportButton_Click);
            // 
            // SaveButton
            // 
            this.SaveButton.Enabled = false;
            this.SaveButton.Location = new System.Drawing.Point(93, 12);
            this.SaveButton.Name = "SaveButton";
            this.SaveButton.Size = new System.Drawing.Size(75, 23);
            this.SaveButton.TabIndex = 1;
            this.SaveButton.Text = "Save";
            this.SaveButton.UseVisualStyleBackColor = true;
            this.SaveButton.Click += new System.EventHandler(this.SaveButton_Click);
            // 
            // CancelButton
            // 
            this.CancelButton.Location = new System.Drawing.Point(174, 12);
            this.CancelButton.Name = "CancelButton";
            this.CancelButton.Size = new System.Drawing.Size(75, 23);
            this.CancelButton.TabIndex = 2;
            this.CancelButton.Text = "Cancel";
            this.CancelButton.UseVisualStyleBackColor = true;
            this.CancelButton.Click += new System.EventHandler(this.CancelButton_Click);
            // 
            // ImportForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(261, 48);
            this.Controls.Add(this.CancelButton);
            this.Controls.Add(this.SaveButton);
            this.Controls.Add(this.ImportButton);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ImportForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Import Excel File";
            this.ResumeLayout(false);



        }

        private Button ImportButton;
        private Button SaveButton;
        private new Button CancelButton;
        private System.Windows.Forms.DataGridView dataGridView1;
    }
}