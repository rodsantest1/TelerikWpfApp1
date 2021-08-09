using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using Telerik.Windows.Documents.Spreadsheet.FormatProviders;
using Telerik.Windows.Documents.Spreadsheet.FormatProviders.OpenXml.Xlsx;
using Telerik.Windows.Documents.Spreadsheet.Model;

namespace TelerikWpfApp1
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

        private void OpenDirectoryButton_Click(object sender, RoutedEventArgs e)
        {
            Process.Start(@"..\Debug");
        }

        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
            BusyIndicator.IsBusy = true;
            OpenDirectoryButton.IsEnabled = false;

            await Task.Run(() =>
            {
                /* Step 1 - Create workbook */
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets.Add();
                worksheet.Columns[3].SetWidth(new ColumnWidth(300, true));
                worksheet.Columns[5].SetWidth(new ColumnWidth(150, true));
                worksheet.Columns[6].SetWidth(new ColumnWidth(150, true));
                worksheet.Columns[7].SetWidth(new ColumnWidth(150, true));

                /* Step 2 - Set value of cell */
                CellSelection selection = worksheet.Cells[1, 1]; //B2 cell 
                selection.SetFormat(new CellValueFormat("\"0\"#"));
                selection.SetValue(0700);

                CellSelection selection2 = worksheet.Cells[1, 2]; //C2 cell 
                selection2.SetFormat(new CellValueFormat("\"0\"#"));
                selection2.SetValue(1);

                CellSelection selection3 = worksheet.Cells[1, 3]; //D3 cell 
                selection3.SetFormat(new CellValueFormat("@"));
                selection3.SetHorizontalAlignment(RadHorizontalAlignment.Center);
                selection3.SetValue("Center this text");

                CellSelection selection4 = worksheet.Cells[1, 4]; //E2 cell 
                selection4.SetFormat(new CellValueFormat("\"0\"#"));
                int test = 1;
                selection4.SetValueAsText(test.ToString());

                CellSelection selection5 = worksheet.Cells[1, 5]; //F2 cell 
                var test2 = DateTime.Now.ToString();
                selection5.SetValue(test2);

                CellSelection selection6 = worksheet.Cells[1, 6]; //G2 cell 
                var test3 = DateTime.Now.ToLongTimeString();
                selection6.SetValue(test3);

                CellSelection selection7 = worksheet.Cells[1, 7]; //H2 cell 
                var test4 = "2021-08-06T16:13:39.8725051-04:00";
                selection7.SetValue(test4);

                CellSelection selection8 = worksheet.Cells[1, 8]; //I2 cell 
                //selection8.SetFormat(new CellValueFormat("\"0\"#"));
                string test5 = "1";
                selection8.SetValue(test5);

                CellSelection selection9 = worksheet.Cells[1, 9]; //J2 cell 
                int test6 = 1;
                selection9.SetValueAsText($"{test6}");

                CellSelection selection10 = worksheet.Cells[1, 10]; //K2 cell 
                //selection10.SetFormat(new CellValueFormat("DateTime"));
                string test7 = "-0:01:57";
                selection10.SetValueAsText($"{test7}");



                /* Step 3 - Export to xlsx */
                string fileName = "SampleFile.xlsx";

                IWorkbookFormatProvider formatProvider = new XlsxFormatProvider();

                using (Stream output = new FileStream(fileName, FileMode.Create))
                {
                    formatProvider.Export(workbook, output);
                }
            });

            BusyIndicator.IsBusy = false;
            StatusLabel.Content = $"Excel document processing completed normally on {DateTime.Now}";
            OpenDirectoryButton.IsEnabled = true;

        }
    }
}
