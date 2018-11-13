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
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelOpener
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private List<string> currentSheetNames = new List<string>();

        public MainWindow()
        {
            InitializeComponent();
            SheetPickerComboBox.ItemsSource = currentSheetNames;
        }

        private void RefreshSheetsList()
        {
            try
            {
                Excel.Application app = new Excel.Application();

                Excel.Workbook wb = app.Workbooks.Open(FilePathTextBox.Text, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing);

                List<string> newSheetNames = new List<string>();

                foreach (Excel.Worksheet ws in wb.Worksheets)
                {
                    newSheetNames.Add(ws.Name);
                }

                if (!currentSheetNames.SequenceEqual(newSheetNames))
                {
                    currentSheetNames = newSheetNames;
                    SheetPickerComboBox.ItemsSource = currentSheetNames;
                }
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                Console.WriteLine("bad file path");
            }
        }

        private void BrowseButton_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog filePicker = new Microsoft.Win32.OpenFileDialog();

            filePicker.DefaultExt = ".xlsx";
            filePicker.Filter = "Excel File (*.xlsx)|*.xlsx";

            Nullable<bool> filePicked = filePicker.ShowDialog();

            if (filePicked == true)
            {
                FilePathTextBox.Text = filePicker.FileName;
                RefreshSheetsList();
            }
        }

        private void RefreshSheetsButton_Click(object sender, RoutedEventArgs e)
        {
            RefreshSheetsList();
        }

        private void LoadButton_Click(object sender, RoutedEventArgs e)
        {
            if (String.IsNullOrEmpty(SheetPickerComboBox.Text))
            {
                Console.WriteLine("No sheet name");
                return;
            }

            try
            {
                Excel.Application app = new Excel.Application();

                Excel.Workbook wb = app.Workbooks.Open(FilePathTextBox.Text, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing);

                Excel.Worksheet sheet = (Excel.Worksheet)wb.Sheets[SheetPickerComboBox.Text];
                Excel.Range excelRange = sheet.UsedRange;

                Console.WriteLine(excelRange.Table().ToString());
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                Console.WriteLine("bad file path");
            }

            



            //foreach (Excel.Range row in excelRange.Rows)
            //{
            //    int rowNumber = row.Row;


            //    string[] A4D4 = GetRange("A" + rowNumber + ":F" + rowNumber + "", sheet);

            //}

            //public string[] GetRange(string range, Worksheet excelWorksheet)
            //{
            //    Microsoft.Office.Interop.Excel.Range workingRangeCells =
            //      excelWorksheet.get_Range(range, Type.Missing);
            //    //workingRangeCells.Select();

            //    System.Array array = (System.Array)workingRangeCells.Cells.Value2;
            //    string[] arrayS = this.ConvertToStringArray(array);

            //    return arrayS;
            //}
        }
    }
}
