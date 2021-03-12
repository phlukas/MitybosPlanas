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
        private List<Recipe> recipes = new List<Recipe>();
        public MainWindow()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            InitializeComponent();
            ReadRecipes("Receptai");
            FillComboBox(comboBox);
            FillComboBox(comboBox1);
            FillComboBox(comboBox2);
            FillComboBox(comboBox3);
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            if(textBox.Text.Length > 1 && textBox.Text.Length <= 50)
            {
                CreateExcel(textBox.Text);
            }
            else
            {
                MessageBox.Show("Blogas plano pavadinimas", "Alert", MessageBoxButton.OK, MessageBoxImage.Error);
            }
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

            MessageBox.Show("Dokumentas sukurtas sėkmingai", "Alert", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void FillComboBox(ComboBox box)
        {
            List<string> options = new List<string>();
            foreach (Recipe recipe in recipes)
            {
                options.Add(recipe.Title);
            }
            box.ItemsSource = options;
        }
        private void ReadRecipes(string folderName)
        {
            string[] paths = Directory.GetFiles(System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, folderName));
            foreach (string path in paths)
            {
                recipes.Add(new Recipe(path));
            }
        }
    }
}
