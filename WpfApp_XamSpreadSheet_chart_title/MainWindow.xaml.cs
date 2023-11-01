using Infragistics.Documents.Excel.Charts;
using Infragistics.Documents.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace WpfApp_XamSpreadSheet_chart_title
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            Workbook workbook = new Workbook(WorkbookFormat.Excel2007);
            Worksheet sheet = workbook.Worksheets.Add("Sheet1");

            // X axis labels
            sheet.GetCell("A1").Value = "January";
            sheet.GetCell("B1").Value = "February";
            sheet.GetCell("C1").Value = "March";
            sheet.GetCell("D1").Value = "April";

            // Data
            sheet.GetCell("A3").Value = 10;
            sheet.GetCell("B3").Value = 20;
            sheet.GetCell("C3").Value = 30;
            sheet.GetCell("D3").Value = 40;

            sheet.GetCell("A5").Value = 15;
            sheet.GetCell("B5").Value = 25;
            sheet.GetCell("C5").Value = 23;
            sheet.GetCell("D5").Value = 45;

            sheet.GetCell("A7").Value = 13;
            sheet.GetCell("B7").Value = 23;
            sheet.GetCell("C7").Value = 39;
            sheet.GetCell("D7").Value = 11;

            WorksheetCell cell1 = sheet.GetCell("E7");
            WorksheetCell cell2 = sheet.GetCell("M30");

            WorksheetChart chart1 = sheet.Shapes.AddChart(Infragistics.Documents.Excel.Charts.ChartType.ColumnClustered, cell1, new Point(0, 0), cell2, new Point(100, 100));
            chart1.SetSourceData("A1:D1,A3:D7", true);

            // Create ChartTitle
            ChartTitle chartTitle = new ChartTitle();

            // Setting ChartTitle.Text
            chartTitle.Text = new Infragistics.Documents.Excel.FormattedString("Title Text");

            // Setting Text Height
            chartTitle.Text.GetFont(0).Height = 500;

            // Setting ChartTitle
            chart1.ChartTitle = chartTitle;

            xamSpreadsheet1.Workbook = workbook;

        }
    }
}
