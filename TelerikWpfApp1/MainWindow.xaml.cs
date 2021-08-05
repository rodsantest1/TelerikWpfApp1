using System;
using System.Diagnostics;
using System.IO;
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

            /* Step 1 - Create workbook */
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets.Add();

            /* Step 2 - Set value of cell */
            CellSelection selection = worksheet.Cells[1, 1]; //B2 cell 
            selection.SetValue(07);

            /* Step 3 - Export to xlsx */
            string fileName = "SampleFile.xlsx";

            IWorkbookFormatProvider formatProvider = new XlsxFormatProvider();

            using (Stream output = new FileStream(fileName, FileMode.Create))
            {
                formatProvider.Export(workbook, output);
            }

            StatusLabel.Content = $"Excel document processing completed normally on {DateTime.Now}";
        }

        private void OpenDirectoryButton_Click(object sender, RoutedEventArgs e)
        {
            Process.Start(@"..\Debug");
        }
    }
}
