using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using OfficeOpenXml;

namespace CSV_Excel
{
    public partial class MainWindow : Window
    {
        public ObservableCollection<ObservableCollection<string>> Data { get; set; }
        public MainWindow()
        {
            InitializeComponent();
            Data = new ObservableCollection<ObservableCollection<string>> { new ObservableCollection<string> { "" } };
            dataGrid.ItemsSource = Data;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            AddDataGridColumns();
        }
        private void AddDataGridColumns()
        {
            dataGrid.Columns.Clear();
            int columnCount = Data.Count > 0 ? Data[0].Count : 0;
            for (int i = 0; i < columnCount; i++)
            {
                var column = new DataGridTextColumn
                {
                    Binding = new System.Windows.Data.Binding($"[{i}]"),
                    Header = "",
                    IsReadOnly = false
                };
                dataGrid.Columns.Add(column);
            }
        }
        private void ExportToCSV_Click(object sender, RoutedEventArgs e)
        {
            var saveFileDialog = new Microsoft.Win32.SaveFileDialog { Filter = "CSV files (*.csv)|*.csv" };
            if (saveFileDialog.ShowDialog() == true)
            {
                using (var writer = new StreamWriter(saveFileDialog.FileName))
                {
                    foreach (var row in Data)
                    {
                        writer.WriteLine(string.Join(",", row));
                    }
                }
            }
        }
        private void ImportFromCSV_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new Microsoft.Win32.OpenFileDialog { Filter = "CSV files (*.csv)|*.csv" };
            if (openFileDialog.ShowDialog() == true)
            {
                using (var reader = new StreamReader(openFileDialog.FileName))
                {
                    Data.Clear();
                    while (!reader.EndOfStream)
                    {
                        var line = reader.ReadLine();
                        var values = line.Split(',');
                        var newRow = new ObservableCollection<string>();
                        foreach (var value in values)
                        {
                            newRow.Add(value);
                        }
                        Data.Add(newRow);
                    }
                    AddDataGridColumns();
                }
            }
        }
        private void ImportFromExcel_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new Microsoft.Win32.OpenFileDialog { Filter = "Excel files (*.xlsx)|*.xlsx" };
            if (openFileDialog.ShowDialog() == true)
            {
                using (var package = new ExcelPackage(new FileInfo(openFileDialog.FileName)))
                {
                    var worksheet = package.Workbook.Worksheets.First();
                    Data.Clear();
                    for (int row = 1; row <= worksheet.Dimension.End.Row; row++)
                    {
                        var newRow = new ObservableCollection<string>();
                        for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                        {
                            var cellValue = worksheet.Cells[row, col].Text;
                            newRow.Add(cellValue);
                        }
                        Data.Add(newRow);
                    }
                }
                AddDataGridColumns();
            }
        }

        private void ExportToExcel_Click(object sender, RoutedEventArgs e)
        {
            var saveFileDialog = new Microsoft.Win32.SaveFileDialog { Filter = "Excel files (*.xlsx)|*.xlsx" };
            if (saveFileDialog.ShowDialog() == true)
            {
                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Sheet1");
                    for (int row = 0; row < Data.Count; row++)
                    {
                        for (int col = 0; col < Data[row].Count; col++)
                        {
                            worksheet.Cells[row + 1, col + 1].Value = Data[row][col];
                        }
                    }
                    package.SaveAs(new FileInfo(saveFileDialog.FileName));
                }
            }
        }
        private void AddRow_Click(object sender, RoutedEventArgs e)
        {
            var newRow = new ObservableCollection<string>(new string[dataGrid.Columns.Count]);
            Data.Add(newRow);
        }
        private void AddColumn_Click(object sender, RoutedEventArgs e)
        {
            foreach (var row in Data)
            {
                row.Add("");
            }
            AddDataGridColumns();
        }
    }
}
