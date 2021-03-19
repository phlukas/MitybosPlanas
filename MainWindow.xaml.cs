using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows;
using OfficeOpenXml;
using OfficeOpenXml.Style;

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
            FillComboBoxes();
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
                FileInfo savePath = new FileInfo(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Mitybos planai", fileName));
                File.Delete(savePath.FullName);
                using var package = new ExcelPackage(savePath);
                var sheet = package.Workbook.Worksheets.Add("Planas");

                sheet.SelectedRange[1, 1, 100, 5].Style.Font.Size = 10;
                sheet.SelectedRange[1, 1, 100, 5].Style.WrapText = true;
                sheet.SelectedRange[1, 1, 100, 5].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                sheet.SelectedRange[1, 1, 100, 5].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                DisplayLogo(sheet);

                sheet.Cells["A2"].Value = "Valgiai";
                sheet.Cells["B2"].Value = "Pirmadienis, Antradienis";
                sheet.Cells["C2"].Value = "Trečiadienis, Ketvirtadienis";
                sheet.Cells["D2"].Value = "Penktadienis, Šeštadienis";
                sheet.Cells["E2"].Value = "Sekmadienis, Pirmadienis";

                sheet.Column(1).Width = 9;
                sheet.Column(2).Width = 18;
                sheet.Column(3).Width = 18;
                sheet.Column(4).Width = 18;
                sheet.Column(5).Width = 18;

                sheet.Cells[2, 1, 2, 5].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                sheet.Cells[2, 1, 2, 5].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
                sheet.Cells[3, 1, 8, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                sheet.Cells[3, 1, 8, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);

                sheet.Cells[2, 1, 8, 5].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                sheet.Cells[2, 1, 8, 5].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                sheet.Cells[2, 1, 8, 5].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                sheet.Cells[2, 1, 8, 5].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;

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

        private void DisplayLogo(ExcelWorksheet sheet)
        {
            FileInfo logo = new FileInfo(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logo.PNG"));
            OfficeOpenXml.Drawing.ExcelPicture pic = sheet.Drawings.AddPicture("logo", logo);
            sheet.SelectedRange[1, 1, 1, 5].Merge = true;
            sheet.Row(1).Height = 55;
            pic.SetSize(50);
            pic.SetPosition(0, 0);
        }

        private void DisplayMeal(ExcelWorksheet sheet)
        {
            try
            {
                List<Recipe> selectedRecipes = new List<Recipe>(12);

                sheet.Cells[3, 1].Value = "Pusryčiai";
                sheet.Cells[5, 1].Value = "Pietūs";
                sheet.Cells[7, 1].Value = "Vakarienė";

                //Pusryčiai
                Recipe recipe = (Recipe)comboBox.SelectedItem;
                AddRecipeIfUnique(selectedRecipes, recipe);

                sheet.Cells[3, 2].Value = recipe.Title;
                sheet.Cells[4, 2].Value = recipe.Ingredients;

                recipe = (Recipe)comboBox1.SelectedItem;
                AddRecipeIfUnique(selectedRecipes, recipe);

                sheet.Cells[3, 3].Value = recipe.Title;
                sheet.Cells[4, 3].Value = recipe.Ingredients;

                recipe = (Recipe)comboBox2.SelectedItem;
                AddRecipeIfUnique(selectedRecipes, recipe);

                sheet.Cells[3, 4].Value = recipe.Title;
                sheet.Cells[4, 4].Value = recipe.Ingredients;

                recipe = (Recipe)comboBox3.SelectedItem;
                AddRecipeIfUnique(selectedRecipes, recipe);

                sheet.Cells[3, 5].Value = recipe.Title;
                sheet.Cells[4, 5].Value = recipe.Ingredients;

                //Pietūs
                recipe = (Recipe)comboBox_Copy.SelectedItem;
                AddRecipeIfUnique(selectedRecipes, recipe);

                sheet.Cells[5, 2].Value = recipe.Title;
                sheet.Cells[6, 2].Value = recipe.Ingredients;

                recipe = (Recipe)comboBox1_Copy.SelectedItem;
                AddRecipeIfUnique(selectedRecipes, recipe);

                sheet.Cells[5, 3].Value = recipe.Title;
                sheet.Cells[6, 3].Value = recipe.Ingredients;

                recipe = (Recipe)comboBox2_Copy.SelectedItem;
                AddRecipeIfUnique(selectedRecipes, recipe);

                sheet.Cells[5, 4].Value = recipe.Title;
                sheet.Cells[6, 4].Value = recipe.Ingredients;

                recipe = (Recipe)comboBox3_Copy.SelectedItem;
                AddRecipeIfUnique(selectedRecipes, recipe);

                sheet.Cells[5, 5].Value = recipe.Title;
                sheet.Cells[6, 5].Value = recipe.Ingredients;

                //Vakarienė
                recipe = (Recipe)comboBox_Copy1.SelectedItem;
                AddRecipeIfUnique(selectedRecipes, recipe);

                sheet.Cells[7, 2].Value = recipe.Title;
                sheet.Cells[8, 2].Value = recipe.Ingredients;

                recipe = (Recipe)comboBox1_Copy1.SelectedItem;
                AddRecipeIfUnique(selectedRecipes, recipe);

                sheet.Cells[7, 3].Value = recipe.Title;
                sheet.Cells[8, 3].Value = recipe.Ingredients;

                recipe = (Recipe)comboBox2_Copy1.SelectedItem;
                AddRecipeIfUnique(selectedRecipes, recipe);

                sheet.Cells[7, 4].Value = recipe.Title;
                sheet.Cells[8, 4].Value = recipe.Ingredients;

                recipe = (Recipe)comboBox3_Copy1.SelectedItem;
                AddRecipeIfUnique(selectedRecipes, recipe);

                sheet.Cells[7, 5].Value = recipe.Title;
                sheet.Cells[8, 5].Value = recipe.Ingredients;

                DisplayRecipes(selectedRecipes, 9, sheet);
            }
            catch (Exception)
            {
                throw;
            }
        }

        private void DisplayRecipes(List<Recipe> selectedRecipes, int row, ExcelWorksheet sheet)
        {
            //string text = InOut.ReadLines(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "UŽKANDŽIAI.txt"));
            //sheet.Cells[row, 1, row, 5].Merge = true;
            //sheet.Cells[row, 1, row, 5].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
            //sheet.Cells[row, 1, row, 5].Value = text;
            //sheet.Row(row).Height = MeasureTextHeight(sheet.Cells[row, 1, row, 5].Text, sheet.Cells[row, 1, row, 5].Style.Font, 5, 0);
            //row++;

            sheet.Cells[row, 1, row, 5].Merge = true;
            sheet.Cells[row, 1, row, 5].Value = "Receptai";
            sheet.Cells[row, 1, row, 5].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
            row++;

            foreach (Recipe recipe in selectedRecipes)
            {
                if (recipe.Description.Length > 1)
                {
                    sheet.Cells[row, 1, row, 5].Merge = true;
                    sheet.Cells[row, 1, row, 5].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                    sheet.Cells[row, 1, row, 5].Value = recipe.Title + " - " + recipe.Description;
                    sheet.Row(row).Height = MeasureTextHeight(sheet.Cells[row, 1, row, 5].Text, sheet.Cells[row, 1, row, 5].Style.Font, 5, 10);
                    row++;
                }
            }
        }

        private double MeasureTextHeight(string text, ExcelFont font, double width, int bonus)
        {
            var bitmap = new Bitmap(1, 1);
            var graphics = Graphics.FromImage(bitmap);

            var pixelWidth = Convert.ToInt32(width * 7);  //7 pixels per excel column width
            var fontSize = font.Size * 1.01f;
            var drawingFont = new Font(font.Name, fontSize);
            var size = graphics.MeasureString(text, drawingFont, pixelWidth, new StringFormat { FormatFlags = StringFormatFlags.MeasureTrailingSpaces });

            //72 DPI and 96 points per inch.  Excel height in points with max of 409 per Excel requirements.
            return Math.Min(Convert.ToDouble(size.Height) * 72 / 96, 409) + bonus;
        }

        private void AddRecipeIfUnique(List<Recipe> list, Recipe recipe)
        {
            if (!list.Contains(recipe))
                list.Add(recipe);
        }

        private void FillComboBoxes()
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
            string[] paths = Directory.GetFiles(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, folderName));
            foreach (string path in paths)
            {
                recipes.Add(new Recipe(path));
            }
        }
    }
}
