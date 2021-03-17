using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
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
            FillComboBox();
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

                sheet.Column(1).Width = 9;
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
                else if (e.Message == "Object reference not set to an instance of an object.")
                {
                    MessageBox.Show("Nepasirinkti visi patiekalai", "Alert", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                else
                {
                    throw;
                }
            }
        }

        private void DisplayMeal(ExcelWorksheet sheet)
        {
            try
            {
                Recipe recipe = (Recipe)comboBox.SelectedItem;

                sheet.Cells[2, 1].Value = "Pusryčiai";
                sheet.Cells[4, 1].Value = "Pietūs";
                sheet.Cells[6, 1].Value = "Vakarienė";

                //Pusryčiai
                sheet.Cells[2, 2].Value = recipe.Title;
                sheet.Cells[3, 2].Value = recipe.Ingredients;

                recipe = (Recipe)comboBox1.SelectedItem;

                sheet.Cells[2, 3].Value = recipe.Title;
                sheet.Cells[3, 3].Value = recipe.Ingredients;

                recipe = (Recipe)comboBox2.SelectedItem;

                sheet.Cells[2, 4].Value = recipe.Title;
                sheet.Cells[3, 4].Value = recipe.Ingredients;

                recipe = (Recipe)comboBox3.SelectedItem;

                sheet.Cells[2, 5].Value = recipe.Title;
                sheet.Cells[3, 5].Value = recipe.Ingredients;

                //Pietūs
                recipe = (Recipe)comboBox_Copy.SelectedItem;
                sheet.Cells[4, 2].Value = recipe.Title;
                sheet.Cells[5, 2].Value = recipe.Ingredients;

                recipe = (Recipe)comboBox1_Copy.SelectedItem;

                sheet.Cells[4, 3].Value = recipe.Title;
                sheet.Cells[5, 3].Value = recipe.Ingredients;

                recipe = (Recipe)comboBox2_Copy.SelectedItem;

                sheet.Cells[4, 4].Value = recipe.Title;
                sheet.Cells[5, 4].Value = recipe.Ingredients;

                recipe = (Recipe)comboBox3_Copy.SelectedItem;

                sheet.Cells[4, 5].Value = recipe.Title;
                sheet.Cells[5, 5].Value = recipe.Ingredients;

                //Vakarienė
                recipe = (Recipe)comboBox_Copy1.SelectedItem;
                sheet.Cells[6, 2].Value = recipe.Title;
                sheet.Cells[7, 2].Value = recipe.Ingredients;

                recipe = (Recipe)comboBox1_Copy1.SelectedItem;

                sheet.Cells[6, 3].Value = recipe.Title;
                sheet.Cells[7, 3].Value = recipe.Ingredients;

                recipe = (Recipe)comboBox2_Copy1.SelectedItem;

                sheet.Cells[6, 4].Value = recipe.Title;
                sheet.Cells[7, 4].Value = recipe.Ingredients;

                recipe = (Recipe)comboBox3_Copy1.SelectedItem;

                sheet.Cells[6, 5].Value = recipe.Title;
                sheet.Cells[7, 5].Value = recipe.Ingredients;
            }
            catch (Exception)
            {
                throw;
            }
        }

        private void FillComboBox()
        {
            comboBox.ItemsSource = recipes;
            comboBox1.ItemsSource = recipes;
            comboBox2.ItemsSource = recipes;
            comboBox3.ItemsSource = recipes;
            comboBox_Copy.ItemsSource = recipes;
            comboBox1_Copy.ItemsSource = recipes;
            comboBox2_Copy.ItemsSource = recipes;
            comboBox3_Copy.ItemsSource = recipes;
            comboBox_Copy1.ItemsSource = recipes;
            comboBox1_Copy1.ItemsSource = recipes;
            comboBox2_Copy1.ItemsSource = recipes;
            comboBox3_Copy1.ItemsSource = recipes;
        }
        private void ReadRecipes(string folderName)
        {
            string[] paths = Directory.GetFiles(System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, folderName));
            foreach (string path in paths)
            {
                recipes.Add(new Recipe(path));
            }
        }

        private Recipe GetRecipe(string title)
        {
            foreach (Recipe recipe in recipes)
            {
                if (recipe.Title == title)
                {
                    return recipe;
                }
            }
            return null;
        }
    }
}
