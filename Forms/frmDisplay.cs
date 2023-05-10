using System;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Windows.Forms;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Workbook = DocumentFormat.OpenXml.Spreadsheet.Workbook;
using Worksheet = DocumentFormat.OpenXml.Spreadsheet.Worksheet;
using Sheets = DocumentFormat.OpenXml.Spreadsheet.Sheets;
using Row = DocumentFormat.OpenXml.Spreadsheet.Row;
using DataTable = System.Data.DataTable;

namespace PreprocessingProject.Forms
{
    public partial class frmDisplay : Form
    {
        DataTable OrignalADGVdt = null;
        DataTable dt;
        public frmDisplay (DataTable dt)
        {
            InitializeComponent();
            this.dt = dt;
        }
        private void frmDisplay_Load(object sender, EventArgs e)
        {
            ADGV.DataSource = dt;
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
        private void btn_Export_Click(object sender, EventArgs e)
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
