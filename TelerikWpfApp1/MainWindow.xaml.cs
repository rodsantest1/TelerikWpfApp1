using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using Telerik.Windows.Documents.Fixed.Model;
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

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            OpenDirectoryButton.IsEnabled = false;
        }

        private async void PdfExportButton_Click(object sender, RoutedEventArgs e)
        {
            BusyIndicator.IsBusy = true;
            OpenDirectoryButton.IsEnabled = false;

            await Task.Run(() =>
            {
                Workbook workbook = CreateWorkbook();

                ExportToPdf(workbook);
            });

            BusyIndicator.IsBusy = false;
            StatusLabel.Content = $"Pdf document processing completed normally on {DateTime.Now}";
            OpenDirectoryButton.IsEnabled = true;
        }

        private async void ExcelExportButton_Click(object sender, RoutedEventArgs e)
        {
            BusyIndicator.IsBusy = true;
            OpenDirectoryButton.IsEnabled = false;

            await Task.Run(() =>
            {
                Workbook workbook = CreateWorkbook();

                ExportToExcel(workbook);
            });

            BusyIndicator.IsBusy = false;
            StatusLabel.Content = $"Excel document processing completed normally on {DateTime.Now}";
            OpenDirectoryButton.IsEnabled = true;
        }

        private static void ExportToPdf(Workbook workbook)
        {
            Telerik.Windows.Documents.Spreadsheet.FormatProviders.Pdf.PdfFormatProvider pdfFormatProvider = new Telerik.Windows.Documents.Spreadsheet.FormatProviders.Pdf.PdfFormatProvider();
            using (Stream output = File.OpenWrite("Sample.pdf"))
            {
                pdfFormatProvider.Export(workbook, output);
            }
        }

        private static void ExportToExcel(Workbook workbook)
        {
            /* Step 3 - Export to xlsx */
            string fileName = "SampleFile.xlsx";

            IWorkbookFormatProvider excelFormatProvider = new XlsxFormatProvider();

            using (Stream output = new FileStream(fileName, FileMode.Create))
            {
                excelFormatProvider.Export(workbook, output);
            }
        }

        private static Workbook CreateWorkbook()
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
            selection4.SetFormat(new CellValueFormat("0#"));
            int test = 10;
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
            string test7 = "-0:01:50";
            selection10.SetValueAsText($"{test7}");

            CellSelection selection11 = worksheet.Cells[1, 11]; //L2 cell 
            decimal test8 = -0.50M;
            selection11.SetValue($"{test8}");
            //Todo: note order between these matters
            selection11.SetFormat(new CellValueFormat("General"));

            CellSelection selection12 = worksheet.Cells[1, 12]; //M2 cell 
            TimeSpan test9 = new TimeSpan(-8, 08, 08);
            selection12.SetValueAsText($"{test9}zzz");

            CellSelection selection13 = worksheet.Cells[1, 13]; //N2 cell 
            string test10 = "-8:08:08";
            TimeSpan.TryParse(test10, out TimeSpan test10Out);
            selection13.SetValueAsText($"{test10Out}");

            CellSelection selection14 = worksheet.Cells[1, 14]; //O2 cell 
            selection14.SetValue(1.3);
            selection14.SetFormat(new CellValueFormat("#.00"));

            CellSelection selection15 = worksheet.Cells[1, 15]; //P2 cell 
            selection15.SetValue(-0.3);
            selection15.SetFormat(new CellValueFormat("#0.00"));

            CellSelection selection16 = worksheet.Cells[1, 16]; //Q2 cell 
            selection16.SetValue(100.3);
            selection16.SetFormat(new CellValueFormat("#0.00"));

            /* Step 3 - Export to PDF */
            worksheet.WorksheetPageSetup.PageOrientation = Telerik.Windows.Documents.Model.PageOrientation.Landscape;
            worksheet.WorksheetPageSetup.FitToPages = true;
            worksheet.WorksheetPageSetup.PrintOptions.PrintGridlines = true;
            return workbook;
        }
    }
}
