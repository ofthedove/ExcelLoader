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
            Excel.Application app = new Excel.Application();

            try
            {
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
            Excel.Application app = new Excel.Application();

            try
            {
                Excel.Workbook wb = app.Workbooks.Open(FilePathTextBox.Text, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing);

                foreach (Excel.Worksheet ws in wb.Worksheets)
                {
                    Console.WriteLine(ws.Name);
                }
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                Console.WriteLine("bad file path");
            }
        }
    }
}
