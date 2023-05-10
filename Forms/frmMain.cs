using System;
using System.ComponentModel;
using System.Data;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System.IO;
using DataTable = System.Data.DataTable;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Globalization;

namespace PreprocessingProject.Forms
{
    public partial class frmMain : Form
    {
        DataTable OrignalADGVdt = null, data = null, dtResults = null, dtOutliers = null, dtNormalizedData = null;
        DataSet dsFrequencies = null, dsFiveNumbers = null;
        List<List<double>> lstOutliers = new List<List<double>>();
        
        string sFileName = "";
        Excel.Application xlApp;
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;
        OleDbConnection MyCon;
        OleDbDataAdapter MyAdptr;  
        public frmMain()
        {
            InitializeComponent();
        }
        private void ADGV_FilterStringChanged(object sender, Zuby.ADGV.AdvancedDataGridView.FilterEventArgs e)
        {
            Zuby.ADGV.AdvancedDataGridView fdgv = ADGV;
            DataTable dt = null;
            if (OrignalADGVdt == null)
            {
                OrignalADGVdt = (DataTable)fdgv.DataSource;
            }
            if (fdgv.FilterString.Length > 0)
            {
                dt = (DataTable)fdgv.DataSource;
            }
            else//Clear Filter
            {
                dt = OrignalADGVdt;
            }

            fdgv.DataSource = dt.Select(fdgv.FilterString).CopyToDataTable();
        }
        private void ADGV_SortStringChanged(object sender, Zuby.ADGV.AdvancedDataGridView.SortEventArgs e)
        {
            if (e.SortString.Length == 0)
            {
                return;
            }
            string[] strtok = e.SortString.Split(',');
            foreach (string str in strtok)
            {
                string[] columnorder = str.Split(']');
                ListSortDirection lds = ListSortDirection.Ascending;
                if (columnorder[1].Trim().Equals("DESC"))
                {
                    lds = ListSortDirection.Descending;
                }
                ADGV.Sort(ADGV.Columns[columnorder[0].Replace('[', ' ').Trim()], lds);
            }
        }
        #region Helper Functions
        private void convertToDouble()
        {
            foreach (DataRow row in data.Rows)
            {
                foreach (DataColumn column in data.Columns)
                {
                    if (row[column.ColumnName].ToString() != "NA")
                        row[column.ColumnName] = Convert.ToDouble(row[column].ToString(), CultureInfo.InvariantCulture);
                }
            }
        }
        private void convertLstOfLstsToDatatable(List<List<double>> listOfLists, DataTable dataTable)
        {
            // Add rows to the DataTable
            int maxCount = 0;
            foreach (List<double> sublist in listOfLists)
            {
                if (sublist.Count > maxCount)
                {
                    maxCount = sublist.Count;
                }
            }

            for (int i = 0; i < maxCount; i++)
            {
                DataRow dr = dataTable.NewRow();
                for (int j = 0; j < listOfLists.Count; j++)
                {
                    if (i < listOfLists[j].Count)
                    {
                        dr[j] = listOfLists[j][i];
                    }
                    else
                    {
                        dr[j] = DBNull.Value;
                    }
                }
                dataTable.Rows.Add(dr);
            }


            //int maxListSize = listOfLists.Max(list => list.Count);
            //for (int i = 0; i < maxListSize; i++)
            //{
            //    DataRow dataRow = dataTable.NewRow();
            //    for (int j = 0; j < listOfLists.LongCount(); j++)
            //    {
            //        List<string> list = listOfLists[j];
            //        string item = i < list.Count ? list[i] : null;
            //        dataRow[j] = item;
            //    }
            //    dataTable.Rows.Add(dataRow);
            //}
        }
        private double GetMean(double[] values)
        {
            return values.Sum() / values.Length;
        }
        private double GetMedian(double[] values)
        {
            Array.Sort(values);
            int middle = values.Length / 2;
            if (values.Length % 2 == 0)
            {
                return (values[middle] + values[middle - 1]) / 2.0;
            }
            else
            {
                return values[middle];
            }
        }
        private double GetMode(double[] values)
        {      
            var groups = values.GroupBy(x => x);
            var modeGroup = groups.OrderByDescending(x => x.Count()).First();
            return modeGroup.Key;
        }
        private double GetVariance(double[] values)
        {
            double mean = values.Sum() / values.Length;
            double sumOfSquares = 0;
            foreach (double value in values)
            {
                sumOfSquares += Math.Pow(value - mean, 2);
            }
            double variance = sumOfSquares / values.Length;
            return variance;
        }
        private double GetQuartile(double[] values, double percentile)
        {
            int index = (int)(percentile * (values.Length - 1));
            double[] sortedData = values.OrderBy(n => n).ToArray();
            double fractional = percentile * (values.Length - 1) - index;
            if (index + 1 < sortedData.Length)
            {
                return sortedData[index] * (1 - fractional) + sortedData[index + 1] * fractional;
            }
            else
            {
                return sortedData[index];
            }
        }
        private double GetIQR(double[] values)
        {
            double q1 = GetQuartile(values, 0.25);
            double q3 = GetQuartile(values, 0.75);
            return q3 - q1;
        }
        private double GetOutlierLowerBound(double[] values)
        {
            double iqr = GetIQR(values);
            double q1 = GetQuartile(values, 0.25);
            return q1 - (1.5 * iqr);
        }
        private double GetOutlierUpperBound(double[] values)
        {
            double iqr = GetIQR(values);
            double q3 = GetQuartile(values, 0.75);
            return q3 + (1.5 * iqr);
        }
        private double[] GetNormalizeMinMax(double[] values)
        {
            double[] normalizedData = new double[values.Length];

            double min = values.Min();
            double max = values.Max();

            for (int i = 0; i < values.Length; i++)
            {
                normalizedData[i] = (values[i] - min) / (max - min);
            }

            return normalizedData;
        }
        #endregion
        private void btnReadFile_Click(object sender, EventArgs e)
        {
            //read csv
            sFileName = "";
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "Select Dataset File";
            ofd.FileName = "";
            ofd.Filter = "CSV File|*.csv;";

            if (ofd.ShowDialog() == DialogResult.OK)
                sFileName = ofd.FileName;

            if (sFileName == "")
                return;

            data = new DataTable();
            using (StreamReader sr = new StreamReader(sFileName))
            {
                using (var parser = new Microsoft.VisualBasic.FileIO.TextFieldParser(sr))
                {
                    parser.TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited;
                    parser.SetDelimiters(",");

                    string[] headers = parser.ReadFields();
                    foreach (string header in headers)
                    {
                        data.Columns.Add(header);
                    }

                    while (!parser.EndOfData)
                    {
                        string[] fields = parser.ReadFields();
                        data.Rows.Add(fields);
                    }
                }
            }

            //count missing data
            int counter = 0;
            foreach (DataRow dr in data.Rows)
            {
                for (int i = 0; i < dr.ItemArray.Length; i++)
                {
                    if (dr[i].ToString() == "NA")
                        counter++;
                }
            }


            //intialize outliers table
            dtOutliers = data.Clone();

            //intialize results table
            dtResults = new DataTable();
            DataColumn newColumn = new DataColumn("Process", typeof(string));
            dtResults.Columns.Add(newColumn);
            foreach (DataColumn col in data.Columns)
            {
                dtResults.Columns.Add(col.ColumnName);
            }

            //intialize Normalized table
            dtNormalizedData = data.Copy();

            //convert data to double
            convertToDouble();

            ADGV.DataSource = data;

            txt_Attributtes.Text = data.Columns.Count.ToString();
            txt_Classes.Text = data.Rows.Count.ToString();
            txt_Missing.Text = counter.ToString();
        }
        private void btnMean_Click(object sender, EventArgs e)
        {
            DataRow newRow = dtResults.NewRow();
            
            foreach (DataColumn col in data.Columns)
            {
                double[] values = data.AsEnumerable()
                    .Select(r => Convert.ToDouble(r[col.ColumnName].ToString()))
                    .ToArray();
                double average = GetMean(values);

                newRow[col.ColumnName] = average.ToString("F5");
            }

            newRow["Process"] = "Mean";
            dtResults.Rows.Add(newRow);

            ADGV_Results.DataSource = dtResults;

            MessageBox.Show("Phase Completed Successfully", "Result", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void btnMedian_Click(object sender, EventArgs e)
        {
            DataRow newRow = dtResults.NewRow();

            foreach (DataColumn col in data.Columns)
            {
                double[] values = data.AsEnumerable()
                    .Select(r => Convert.ToDouble(r[col.ColumnName].ToString()))
                    .ToArray();
                double median = GetMedian(values);

                newRow[col.ColumnName] = median.ToString("F5");
            }

            newRow["Process"] = "Median";
            dtResults.Rows.Add(newRow);

            ADGV_Results.DataSource = dtResults;

            MessageBox.Show("Phase Completed Successfully", "Result", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void btnMode_Click(object sender, EventArgs e)
        {
            DataRow newRow = dtResults.NewRow();

            foreach (DataColumn col in data.Columns)
            {
                double[] values = data.AsEnumerable()
                    .Select(r => Convert.ToDouble(r[col.ColumnName].ToString()))
                    .ToArray();
                double mode = GetMode(values);

                newRow[col.ColumnName] = mode.ToString("F5");
            }

            newRow["Process"] = "Mode";
            dtResults.Rows.Add(newRow);

            ADGV_Results.DataSource = dtResults;

            MessageBox.Show("Phase Completed Successfully", "Result", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void btnVariance_Click(object sender, EventArgs e)
        {
            DataRow newRowV = dtResults.NewRow();
            DataRow newRowSD = dtResults.NewRow();

            foreach (DataColumn col in data.Columns)
            {
                double[] values = data.AsEnumerable()
                    .Select(r => Convert.ToDouble(r[col.ColumnName].ToString()))
                    .ToArray();
                double variance = GetVariance(values);

                newRowV[col.ColumnName] = variance.ToString("F5");

                //calculate SD
                newRowSD[col.ColumnName] = Math.Sqrt(variance).ToString("F5");
            }

            newRowV["Process"] = "Variance";
            newRowSD["Process"] = "Standard Deviation";

            dtResults.Rows.Add(newRowV);
            dtResults.Rows.Add(newRowSD);

            ADGV_Results.DataSource = dtResults;

            MessageBox.Show("Phase Completed Successfully", "Result", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void btnIQR_Click(object sender, EventArgs e)
        {
            DataRow newRow = dtResults.NewRow();

            foreach (DataColumn col in data.Columns)
            {
                double[] values = data.AsEnumerable()
                    .Select(r => Convert.ToDouble(r[col.ColumnName].ToString()))
                    .ToArray();
                double iqr = GetIQR(values);

                newRow[col.ColumnName] = iqr.ToString("F5");
            }

            newRow["Process"] = "IQR";
            dtResults.Rows.Add(newRow);

            ADGV_Results.DataSource = dtResults;

            MessageBox.Show("Phase Completed Successfully", "Result", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void btnExportResults_Click(object sender, EventArgs e)
        {
            //install Open-XML-SDK by microsoft v2.9.1 from NuGet
            if (ADGV_Results.Rows.Count > 0)
            {
                FileStream stream = new FileStream("ExportData.xlsx", FileMode.Create);

                SpreadsheetDocument spreadsheetdocument = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
                WorkbookPart workbookpart = spreadsheetdocument.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();
                WorksheetPart worksheetpart = workbookpart.AddNewPart<WorksheetPart>();
                worksheetpart.Worksheet = new Worksheet(new SheetData());
                Sheets sheets = spreadsheetdocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());
                Sheet sheet = new Sheet()
                {
                    Id = spreadsheetdocument.WorkbookPart.GetIdOfPart(worksheetpart),
                    SheetId = 1,
                    Name = "Exported Data"
                };
                sheets.Append(sheet);
                Worksheet worksheet = worksheetpart.Worksheet;
                SheetData sheetData = worksheet.GetFirstChild<SheetData>();


                Row header = new Row();
                for (int i = 0; i < ADGV_Results.Columns.Count; i++)
                {
                    Cell c = new Cell()
                    {
                        CellValue = new CellValue(ADGV_Results.Columns[i].HeaderText.ToString()),
                        DataType = CellValues.String
                    };
                    header.Append(c);
                }
                sheetData.Append(header);

                foreach (DataGridViewRow dgvr in ADGV_Results.Rows)
                {
                    Row r = new Row();
                    for (int i = 0; i < dgvr.Cells.Count; i++)
                    {
                        Cell c = new Cell()
                        {
                            CellValue = new CellValue(dgvr.Cells[i].Value.ToString()),
                            DataType = CellValues.String
                        };
                        r.Append(c);
                    }
                    sheetData.Append(r);
                }
                worksheetpart.Worksheet.Save();
                spreadsheetdocument.Close();
                stream.Close();
                System.Diagnostics.Process.Start("ExportData.xlsx");
            }
        }
        private void btnNormalize_Click(object sender, EventArgs e)
        {
            foreach (DataColumn col in data.Columns)
            {
                double[] values = data.AsEnumerable()
                    .Select(r => Convert.ToDouble(r[col.ColumnName].ToString()))
                    .ToArray();
                double[] normalizedData = GetNormalizeMinMax(values);

                foreach (DataRow row in dtNormalizedData.Rows)
                {
                    int columnIndex = dtNormalizedData.Columns.IndexOf(col.ColumnName);
                    row[columnIndex] = normalizedData[row.Table.Rows.IndexOf(row)];
                }
            }

            frmDisplay frm = new frmDisplay(dtNormalizedData);
            frm.Text = "Normalized Data";
            frm.Show();

            MessageBox.Show("Phase Completed Successfully", "Result", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void btnFiveNumbers_Click(object sender, EventArgs e)
        {
            DataRow newRowMin = dtResults.NewRow();
            DataRow newRowMax = dtResults.NewRow();
            DataRow newRowQ1 = dtResults.NewRow();
            DataRow newRowQ3 = dtResults.NewRow();

            foreach (DataColumn col in data.Columns)
            {
                double[] values = data.AsEnumerable()
                    .Select(r => Convert.ToDouble(r[col.ColumnName].ToString()))
                    .ToArray();

                double q1 = GetQuartile(values, 0.25);
                double q3 = GetQuartile(values, 0.75);

                newRowMin[col.ColumnName] = values.Min().ToString("F5");
                newRowMax[col.ColumnName] = values.Max().ToString("F5");
                newRowQ1[col.ColumnName] = q1.ToString("F5");
                newRowQ3[col.ColumnName] = q3.ToString("F5");
            }

            newRowMin["Process"] = "Min";
            newRowMax["Process"] = "Max";
            newRowQ1["Process"] = "Q1";
            newRowQ3["Process"] = "Q3";

            dtResults.Rows.Add(newRowMin);
            dtResults.Rows.Add(newRowMax);
            dtResults.Rows.Add(newRowQ1);
            dtResults.Rows.Add(newRowQ3);

            ADGV_Results.DataSource = dtResults;

            MessageBox.Show("Phase Completed Successfully", "Result", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void btnBoxPlot_Click(object sender, EventArgs e)
        {
            dsFiveNumbers = new DataSet();
            foreach (DataColumn col in data.Columns)
            {
                double[] values = data.AsEnumerable()
                    .Select(r => Convert.ToDouble(r[col.ColumnName].ToString()))
                    .ToArray();

                double min = values.Min();
                double max = values.Max();
                double q1 = GetQuartile(values, 0.25);               
                double q3 = GetQuartile(values, 0.75);
                double median = GetMedian(values);

                DataTable dtTemp = new DataTable(col.ColumnName);
                dtTemp.Columns.Add("min", typeof(double));
                dtTemp.Columns.Add("q1", typeof(double));
                dtTemp.Columns.Add("median", typeof(double));
                dtTemp.Columns.Add("q3", typeof(double));
                dtTemp.Columns.Add("max", typeof(double));

                DataRow dr = dtTemp.NewRow();
                dr["min"] = min; dr["max"] = max; dr["q1"] = q1; dr["q3"] = q3; dr["median"] = median;
                dtTemp.Rows.Add(dr);

                dsFiveNumbers.Tables.Add(dtTemp);

            }

            frmCharts frm = new frmCharts(dsFiveNumbers, true);
            frm.Text = "Five Number Summaries";
            frm.Show();
        }
        private void btnHistogram_Click(object sender, EventArgs e)
        {
            dsFrequencies = new DataSet();

            foreach (DataColumn column in data.Columns)
            {
                DataTable dtTemp = new DataTable(column.ColumnName);
                dtTemp.Columns.Add("Value", typeof(double));
                dtTemp.Columns.Add("Frequency", typeof(int));

                var groups = data.AsEnumerable().GroupBy(row => row[column]);

                foreach (var group in groups)
                {
                    dtTemp.Rows.Add(group.Key, group.Count());
                }

                dsFrequencies.Tables.Add(dtTemp);
            }

            frmCharts frm = new frmCharts(dsFrequencies, false);
            frm.Show();
        }
        private void btnOutliers_Click(object sender, EventArgs e)
        {
            DataRow newRowL = dtResults.NewRow();
            DataRow newRowU = dtResults.NewRow();

            foreach (DataColumn col in data.Columns)
            {
                double[] values = data.AsEnumerable()
                    .Select(r => Convert.ToDouble(r[col.ColumnName].ToString()))
                    .ToArray();

                double outlierLowerBound = GetOutlierLowerBound(values);
                double outlierUpperBound = GetOutlierUpperBound(values);

                List<double> outliers = data.AsEnumerable()
                    .Where(r => {
                        double value = Convert.ToDouble(r[col.ColumnName].ToString());
                        return value < outlierLowerBound || value > outlierUpperBound;
                    })
                    .Select(row => Convert.ToDouble(row[col.ColumnName]))
                    .ToList();

                lstOutliers.Add(outliers);

                newRowL[col.ColumnName] = outlierLowerBound.ToString("F5");
                newRowU[col.ColumnName] = outlierUpperBound.ToString("F5");
            }

            convertLstOfLstsToDatatable(lstOutliers, dtOutliers);

            
            frmDisplay frm = new frmDisplay(dtOutliers);
            frm.Text = "Outliers Values";
            frm.Show();

            newRowL["Process"] = "Outlier LB";
            newRowU["Process"] = "Outlier UB";

            dtResults.Rows.Add(newRowL);
            dtResults.Rows.Add(newRowU);

            ADGV_Results.DataSource = dtResults;

            MessageBox.Show("Phase Completed Successfully", "Result", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void btnDataCompletion_Click(object sender, EventArgs e)
        {
            // Mean/Median imputation
            foreach (DataColumn col in data.Columns)
            {
                double[] values = data.AsEnumerable()
                    .Where(r => r[col.ColumnName].ToString() != "NA")
                    .Select(r => Convert.ToDouble(r[col.ColumnName].ToString()))
                    .ToArray();
                double imputedValue = GetMean(values); // we can use GetMedian(values)
                foreach (DataRow row in data.Rows)
                {
                    if (row[col.ColumnName].ToString() == "NA")
                    {
                        row[col.ColumnName] = imputedValue.ToString("F5");
                    }
                }
            }

            MessageBox.Show("Phase Completed Successfully", "Result", MessageBoxButtons.OK, MessageBoxIcon.Information);

            ADGV.Refresh();
        }
        private void btnExportNewData_Click(object sender, EventArgs e)
        {
            //install Open-XML-SDK by microsoft v2.9.1 from NuGet
            if (ADGV.Rows.Count > 0)
            {
                FileStream stream = new FileStream("ExportData.xlsx", FileMode.Create);

                SpreadsheetDocument spreadsheetdocument = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
                WorkbookPart workbookpart = spreadsheetdocument.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();
                WorksheetPart worksheetpart = workbookpart.AddNewPart<WorksheetPart>();
                worksheetpart.Worksheet = new Worksheet(new SheetData());
                Sheets sheets = spreadsheetdocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());
                Sheet sheet = new Sheet()
                {
                    Id = spreadsheetdocument.WorkbookPart.GetIdOfPart(worksheetpart),
                    SheetId = 1,
                    Name = "Exported Data"
                };
                sheets.Append(sheet);
                Worksheet worksheet = worksheetpart.Worksheet;
                SheetData sheetData = worksheet.GetFirstChild<SheetData>();


                Row header = new Row();
                for (int i = 0; i < ADGV.Columns.Count; i++)
                {
                    Cell c = new Cell()
                    {
                        CellValue = new CellValue(ADGV.Columns[i].HeaderText.ToString()),
                        DataType = CellValues.String
                    };
                    header.Append(c);
                }
                sheetData.Append(header);

                foreach (DataGridViewRow dgvr in ADGV.Rows)
                {
                    Row r = new Row();
                    for (int i = 0; i < dgvr.Cells.Count; i++)
                    {
                        Cell c = new Cell()
                        {
                            CellValue = new CellValue(dgvr.Cells[i].Value.ToString()),
                            DataType = CellValues.String
                        };
                        r.Append(c);
                    }
                    sheetData.Append(r);
                }
                worksheetpart.Worksheet.Save();
                spreadsheetdocument.Close();
                stream.Close();
                System.Diagnostics.Process.Start("ExportData.xlsx");
            }
        }
    }
}
