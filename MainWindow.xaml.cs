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
            CreateExcel("Testas");
        }

        private void CreateExcel(string fileName)
        {
            string savePath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileName);

            
        }
    }
}
