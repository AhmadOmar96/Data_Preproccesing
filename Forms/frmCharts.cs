using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using Series = System.Windows.Forms.DataVisualization.Charting.Series;
using System.Windows.Forms.DataVisualization.Charting;
using System.Reflection;
using System.Linq;
using DataTable = System.Data.DataTable;
using Color = System.Drawing.Color;

namespace PreprocessingProject.Forms
{
    public partial class frmCharts : Form
    {
        DataSet ds;
        bool isBoxPlot = false;
        public frmCharts(DataSet ds, bool isBoxPlot)
        {
            InitializeComponent();
            this.ds = ds;
            this.isBoxPlot = isBoxPlot;
        }
        private void HistoCharts_Load(object sender, EventArgs e)
        {
            //make color list from the system colors
            List<Color> colorList = new List<Color>();
            PropertyInfo[] colorProperties = typeof(Color).GetProperties(BindingFlags.Static | BindingFlags.Public);
            foreach (PropertyInfo property in colorProperties)
            {
                if (property.PropertyType == typeof(Color))
                {
                    colorList.Add((Color)property.GetValue(null));
                }
            }

            int attributtesCount = ds.Tables.Count;//to know how many chart we will create
            int colorCounter = 20;//to start with dark colors

            for (int i = 0; i < attributtesCount; i++)
            {
                ChartArea chrt = new ChartArea();
                chrt.Name = "ChartArea" + i;
                histoChart.ChartAreas.Add(chrt);
                Series s = new Series(ds.Tables[i].TableName);
                s.ChartArea = "ChartArea" + i;               
                s.Legend = "Legend1";
                s.Color = colorList[i+colorCounter];
                if (isBoxPlot)
                {
                    s.ChartType = SeriesChartType.BoxPlot;
                }
                else
                {
                    s.ChartType = SeriesChartType.SplineArea;
                }
                
                histoChart.Series.Add(s);
            }

            //Fill charts
            if (isBoxPlot)
            {
                foreach (Series s in histoChart.Series)
                {
                    double[] FiveNumbersValues = ds.Tables[s.Name].Rows[0].ItemArray.Select(x => Convert.ToDouble(x)).ToArray();
                    s.Points.Add(FiveNumbersValues);
                }
            }
            else
            {
                foreach (DataTable dt in ds.Tables)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        
                        histoChart.Series[dt.TableName].Points.AddXY(dr[0], dr[1]);
                    }
                }
            }                      
        }
    }
}
