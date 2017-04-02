using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Data.OleDb;
using System.Data;

namespace Excolidator
{
    public partial class MainForm : Form
    {
        //To keep count of the number of columns in the master sheet
        int masterColCount = 0;
        //To record the name of the activesheet in the Master Workbook
        string masterSheet;
        //To store the master workbook
        string masterWorkbook = null;
        //To store the collection of slave sheets
        List<string> slaveFileNames = new List<string>();
        List<string> slaveNames = new List<string>();
        string masterText = "Upload the master workbook here.";
        //To store the index of the unique column in master workbook
        int masterIndex = 0;

        public MainForm()
        {
            InitializeComponent();
        }

        //To save the Master Workbook
        private void masterButton_Click(object sender, EventArgs e)
        {
            if (slaveFileNames.Count != 0)
            {
                slaveFileNames.Clear();
                slaveNames.Clear();
            }
            slaveListBox.DataSource = null;
            masterIndex = 0;
            masterWorkbook = null;
            masterColCount = 0;
            masterSheet = null;
            masterTextBox.Text = "The selected workbook is being analysed...Please wait.";
            OpenFileDialog opMaster = new OpenFileDialog();
            opMaster.Multiselect = false;
            opMaster.ShowDialog();
            if (opMaster.FileNames.Length == 0)
            {
                MessageBox.Show("No Workbook was selected.\nPlease try again...");
                masterTextBox.Text = masterText;
            }
            else
            {
                masterIndex = 0;
                masterWorkbook = null;
                masterColCount = 0;
                masterTextBox.Text = "The selected workbook is being analysed...Please wait.";
                if (opMaster.FileName.Substring(opMaster.FileName.Length - 5) != ".xlsx")
                {
                    MessageBox.Show("The selected Workbook is not an Excel Workbook!!! Please try again...");
                    masterTextBox.Text = masterText;
                }
                else
                {
                    Excel.Application xlApp = null;
                    Excel.Workbooks workbooks = null;
                    Excel.Workbook workbook = null;
                    Excel.Worksheet worksheet = null;
                    Excel.Range usedRange = null;
                    Excel.Range specialCellsRange = null;
                    Excel.Range columnsRange = null;
                    try
                    {
                        xlApp = new Excel.Application();
                        int colCount = 0;
                        masterColCount = 0;
                        workbooks = xlApp.Workbooks;
                        workbook = workbooks.Open(opMaster.FileName);
                        worksheet = workbook.ActiveSheet;
                        masterSheet = worksheet.Name;
                        usedRange= worksheet.UsedRange;
                        specialCellsRange = usedRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible, Type.Missing);
                        columnsRange = specialCellsRange.Columns;
                        colCount = columnsRange.Count;
                        for (int col = 1; col <= colCount; col++)
                        {
                            Excel.Range range = worksheet.Cells[1, col];
                            if (range.Value != null)
                            {
                                masterColCount++;
                                if (range != null)
                                    Marshal.ReleaseComObject(range);
                                continue;
                            }
                            else
                            {
                                if (range != null)
                                    Marshal.ReleaseComObject(range);
                                break;
                            }
                        }
                        workbook.Close();
                        xlApp.Quit();
                        while (masterIndex <= 0 || masterIndex > masterColCount)
                        {
                            try
                            {
                                masterIndex = int.Parse(Showdialog());
                            }
                            catch (FormatException ex)
                            {
                                MessageBox.Show("Enter a valid Integer...");
                            }
                        }
                        MessageBox.Show("The Master Workbook has been successfully added...");
                        masterTextBox.Text = masterWorkbook = opMaster.FileName;
                    }
                    catch (Exception ex)
                    {
                        masterTextBox.Text = masterText;
                        MessageBox.Show("The selected Workbook cannot be opened!!! Please try again...");
                    }
                    finally
                    {
                        if (usedRange != null) Marshal.ReleaseComObject(usedRange);
                        if (specialCellsRange != null) Marshal.ReleaseComObject(specialCellsRange);
                        if (columnsRange != null) Marshal.ReleaseComObject(columnsRange);
                        if (worksheet != null) Marshal.ReleaseComObject(worksheet);
                        if (workbook != null) Marshal.ReleaseComObject(workbook);
                        if (workbooks != null) Marshal.ReleaseComObject(workbooks);
                        if (xlApp != null) Marshal.ReleaseComObject(xlApp);
                    }
                }
            }
        }

        //To obtain the index of the master column
        private string Showdialog()
        {
            Form indexForm = new Form();
            indexForm.Size = new System.Drawing.Size(300, 130);
            indexForm.MinimizeBox = false;
            indexForm.MaximizeBox = false;
            indexForm.StartPosition = FormStartPosition.CenterParent;
            indexForm.FormBorderStyle = FormBorderStyle.FixedSingle;
            indexForm.Text = "Enter Master Column Index";

            Label label = new Label();
            label.Text = "DocSci ID Index starting from 1:";
            label.Top = 20;
            label.Left = 17;
            label.Width = 250;
            label.Height = 30;
            label.BorderStyle = BorderStyle.Fixed3D;
            label.TextAlign = ContentAlignment.MiddleLeft;

            TextBox indexText = new TextBox();
            indexText.Width = 50;
            indexText.Top = 4;
            indexText.Left = 183;
            indexForm.Controls.Add(indexText);
            label.Controls.Add(indexText);
            indexForm.Controls.Add(label);

            Button acceptButton = new Button();
            acceptButton.Text = "OK";
            acceptButton.Left = 110;
            acceptButton.Top = 60;
            acceptButton.DialogResult = DialogResult.OK;
            indexForm.Controls.Add(acceptButton);
            indexForm.AcceptButton = acceptButton;

            if (indexForm.ShowDialog(this) == DialogResult.OK)
                return indexText.Text;
            else return null;
        }

        //To choose the Slave workbooks
        private void slaveButton_Click(object sender, EventArgs e)
        {
            if (masterWorkbook == null)
            {
                MessageBox.Show("Please select a Master Workbook to proceed..");
            }
            else
            {
                OpenFileDialog opSlave = new OpenFileDialog();
                opSlave.Multiselect = true;
                opSlave.ShowDialog();
                if (opSlave.FileNames.Length == 0)
                    MessageBox.Show("No workbooks were selected.\nPlease try again...");
                else
                {
                    slaveListBox.DataSource = null;
                    slaveListBox.DataSource = new List<string>() { "The selected workbooks are being analysed...Please wait." };
                    SlaveSheetsChecker(opSlave.FileNames);
                }
            }
        }

        //For verifying the slave sheets
        private void SlaveSheetsChecker(string[] fileNames)
        {
            List<string> errorFiles = new List<string>();
            List<string> unmatchCountFiles = new List<string>();
            int colCount = 0;
            int slaveCount = 0;
            //Loop to check if the selected files are Excel sheets
            foreach (string fname in fileNames)
            {
                string temp = fname.Substring(fname.LastIndexOf("\\") + 1);
                if (slaveNames.Contains(temp) || masterWorkbook.Substring(masterWorkbook.LastIndexOf("\\") + 1) == temp)
                    continue;
                else if (fname.Substring(fname.Length - 5) != ".xlsx")
                {
                    errorFiles.Add(fname);
                }
                else
                {
                    Excel.Application xlApp = null;
                    Excel.Workbooks workbooks = null;
                    Excel.Workbook workbook = null;
                    Excel.Worksheet worksheet = null;
                    Excel.Range usedRange = null;
                    Excel.Range specialCellsRange = null;
                    Excel.Range columnsRange = null;
                    try
                    {
                        xlApp = new Excel.Application();
                        workbooks = xlApp.Workbooks;
                        workbook = workbooks.Open(fname);
                        worksheet = workbook.Worksheets[masterSheet];
                        usedRange = worksheet.UsedRange;
                        specialCellsRange = usedRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible, Type.Missing);
                        columnsRange = specialCellsRange.Columns;
                        colCount = columnsRange.Count;
                        for (int col = 1; col <= colCount; col++)
                        {
                            Excel.Range range = worksheet.Cells[1, col];
                            if (range.Value != null)
                            {
                                slaveCount++;
                                if (range != null) Marshal.ReleaseComObject(range);
                                continue;
                            }
                            else
                                if (range != null) Marshal.ReleaseComObject(range);
                                break;
                        }
                        if (slaveCount != masterColCount)
                            unmatchCountFiles.Add(fname);
                        else
                        {
                            slaveFileNames.Add(fname);
                            slaveNames.Add(temp);
                        }
                        colCount = 0;
                        slaveCount = 0;
                        workbook.Close();
                        xlApp.Quit();
                    }
                    catch (Exception ex)
                    {
                        errorFiles.Add(fname);
                        continue;
                    }
                    finally
                    {
                        if (usedRange != null) Marshal.ReleaseComObject(usedRange);
                        if (specialCellsRange != null) Marshal.ReleaseComObject(specialCellsRange);
                        if (columnsRange != null) Marshal.ReleaseComObject(columnsRange);
                        if (worksheet != null) Marshal.ReleaseComObject(worksheet);
                        if (workbook != null) Marshal.ReleaseComObject(workbook);
                        if (workbooks != null) Marshal.ReleaseComObject(workbooks);
                        if (xlApp != null) Marshal.ReleaseComObject(xlApp);
                    }
                }
            }
            StringBuilder sb = new StringBuilder();
            if (errorFiles.Count >= 1)
            {
                sb.Append("The following files could not be selected:\n");
                foreach (string file in errorFiles)
                {
                    sb.Append(file + "\n");
                }
            }
            if (unmatchCountFiles.Count >= 1)
            {
                sb.Append(string.Format("\nThe following Workbooks do not match the column count of {0}\n", masterColCount));
                foreach (string files in unmatchCountFiles)
                {
                    sb.Append(files + "\n");
                }
            }
            if (errorFiles.Count == 0 && unmatchCountFiles.Count == 0)
                sb.Append("The workbooks were successfully added...");
            MessageBox.Show(sb.ToString());
            slaveListBox.DataSource = null;
            slaveListBox.DataSource = slaveFileNames;
            errorFiles.Clear();
            unmatchCountFiles.Clear();
        }

        //To clear a specific Slave workbook for the collection
        private void clearButton_Click(object sender, EventArgs e)
        {
            try
            {
                string file = (string)slaveListBox.SelectedItem;
                slaveFileNames.Remove(file);
                slaveNames.Remove(file.Substring(file.LastIndexOf("\\") + 1));
            }
            catch (Exception ex) { }
            finally
            {
                slaveListBox.DataSource = null;
                slaveListBox.DataSource = slaveFileNames;
            }
        }

        //To clear all the workbooks from the slave workbooks list
        private void clearAllButton_Click(object sender, EventArgs e)
        {
            if (slaveFileNames.Count != 0)
            {
                slaveFileNames.Clear();
                slaveNames.Clear();
            }
            slaveListBox.DataSource = null;
        }

        //To consolidate the slave workbooks with the Master Workbook
        private void consolidateButton_Click(object sender, EventArgs e)
        {
            if (masterWorkbook == null && slaveFileNames.Count == 0)
            {
                MessageBox.Show("Select a Master Workbook and the Slave workbooks to proceed...");
                
            }
            else if (masterWorkbook == null)
            {
                MessageBox.Show("Select a Master Workbook to proceed...");
                
            }
            else if (slaveFileNames.Count == 0)
            {
                MessageBox.Show("Select the Slave workbooks to proceed...");
                
            }
            else
            {
                //To hold the individual sheets
                DataSet ds = new DataSet();
                DataTable workbookTable = new DataTable();
                OleDbDataAdapter xlAdapter = new OleDbDataAdapter();
                OleDbConnection conn = new OleDbConnection();
                OleDbCommand comm = new OleDbCommand();

                //To load the individual sheets into the Dataset
                foreach (var file in slaveFileNames)
                {
                    conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + file +
                                            ";Extended Properties=\"Excel 12.0;HDR=YES;\"";
                    conn.Open();
                    comm.Connection = conn;

                    workbookTable.Clear();
                    workbookTable = comm.Connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    comm.CommandText = "SELECT * FROM [" + masterSheet + "$]";
                    xlAdapter.SelectCommand = comm;
                    xlAdapter.Fill(ds);
                    conn.Close();
                }

                conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + masterWorkbook +
                                            ";Extended Properties=\"Excel 12.0;HDR=YES;\"";

                //To update the Master Workbook row by row
                foreach (DataRow row in ds.Tables[0].Rows)
                {
                    comm.CommandText = "UPDATE [" + masterSheet + "$] SET ";
                    for (int i = 0; i < masterColCount; i++)
                    {
                        if (i == masterIndex - 1)
                            continue;

                        if (i == masterColCount - 1)
                            comm.CommandText += ds.Tables[0].Columns[i].ColumnName + " = '" + row[i].ToString() + "'" +
                                " WHERE " + ds.Tables[0].Columns[masterIndex - 1].ToString() + "= '" + row[masterIndex - 1].ToString() + "'";
                        else
                            comm.CommandText += ds.Tables[0].Columns[i].ColumnName + " = '" + row[i].ToString() + "',";
                    }
                    comm.CommandType = CommandType.Text;
                    conn.Open();
                    comm.ExecuteNonQuery();
                    conn.Close();
                }
                MessageBox.Show("The individual workbooks have been consolidated successfully...");
            }
        }
    }
}