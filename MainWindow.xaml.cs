using System;
using System.Collections.Generic;
using System.IO;
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
using OfficeOpenXml;

namespace MitybosPlanas
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            InitializeComponent();

        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            CreateExcel("Testas.xlsx");
        }

        private void CreateExcel(string fileName)
        {
            FileInfo savePath = new FileInfo(System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileName));
            File.Delete(savePath.FullName);
            using var package = new ExcelPackage(savePath);
            var sheet = package.Workbook.Worksheets.Add("Planas");

            sheet.Cells["A1"].Value = "Valgiai";
            sheet.Cells["B1"].Value = "Pirmadienis, Antradienis";
            sheet.Cells["C1"].Value = "Trečiadienis, Ketvirtadienis";
            sheet.Cells["D1"].Value = "Penktadienis, Šeštadienis";
            sheet.Cells["E1"].Value = "Sekmadienis, Pirmadienis";

            sheet.Column(1).Width = 10;

            package.Save();
        }

        
    }
}
