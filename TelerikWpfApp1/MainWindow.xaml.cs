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

                /* Step 2 - Set value of cell */
                CellSelection selection = worksheet.Cells[1, 1]; //B2 cell 
                selection.SetValue(0700);
                selection.SetFormat(new CellValueFormat("\"0\"#"));

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
