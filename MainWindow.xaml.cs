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
            if(textBox.Text.Length > 0 && textBox.Text.Length <= 50)
            {
                CreateExcel(textBox.Text + ".xlsx");
            }
            else
            {
                MessageBox.Show("Blogas plano pavadinimas", "Alert", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CreateExcel(string fileName)
        {
            try
            {
                FileInfo savePath = new FileInfo(System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Mitybos planai", fileName));
                File.Delete(savePath.FullName);
                using var package = new ExcelPackage(savePath);
                var sheet = package.Workbook.Worksheets.Add("Planas");

                sheet.SelectedRange[1, 1, 100, 5].Style.Font.Size = 10;
                sheet.SelectedRange[1, 1, 100, 5].Style.WrapText = true;
                sheet.SelectedRange[1, 1, 100, 5].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                sheet.SelectedRange[1, 1, 100, 5].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                sheet.Cells["A1"].Value = "Valgiai";
                sheet.Cells["B1"].Value = "Pirmadienis, Antradienis";
                sheet.Cells["C1"].Value = "Trečiadienis, Ketvirtadienis";
                sheet.Cells["D1"].Value = "Penktadienis, Šeštadienis";
                sheet.Cells["E1"].Value = "Sekmadienis, Pirmadienis";

                sheet.Column(1).Width = 8;
                sheet.Column(2).Width = 18;
                sheet.Column(3).Width = 18;
                sheet.Column(4).Width = 18;
                sheet.Column(5).Width = 18;

                sheet.Cells[1, 1, 1, 5].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                sheet.Cells[1, 1, 1, 5].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
                sheet.Cells[2, 1, 7, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                sheet.Cells[2, 1, 7, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);

                sheet.Cells[1, 1, 7, 5].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                sheet.Cells[1, 1, 7, 5].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                sheet.Cells[1, 1, 7, 5].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                sheet.Cells[1, 1, 7, 5].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;

                DisplayMeal(sheet);

                package.Save();

                MessageBox.Show("Dokumentas sukurtas sėkmingai", "Alert", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception e)
            {
                if (e.Source == "System.IO.FileSystem")
                {
                    MessageBox.Show("Uždarykite excel dokumentą", "Alert", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                else
                {
                    throw;
                }
            }
        }

        private void DisplayMeal(ExcelWorksheet sheet)
        {

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
