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

namespace ExcelOpener
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

        private void BrowseButton_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog filePicker = new Microsoft.Win32.OpenFileDialog();

            filePicker.DefaultExt = ".xlsx";
            filePicker.Filter = "Excel File (*.xlsx)|*.xlsx";

            Nullable<bool> filePicked = filePicker.ShowDialog();

            if (filePicked == true)
            {
                FilePathTextBox.Text = filePicker.FileName;
            }
        }

        private void LoadButton_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
