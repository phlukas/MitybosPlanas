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
        private List<Recipe> recipesPav = new List<Recipe>();
        private List<Recipe> recipesPries = new List<Recipe>();

        public MainWindow()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            InitializeComponent();
            ReadRecipes(@"Receptai\Pusryčiai", recipes, "Pusryčiai");
            ReadRecipes(@"Receptai\Pietūs", recipes, "Pietūs");
            ReadRecipes(@"Receptai\Vakarienė", recipes, "Vakarienė");
            ReadRecipes(@"Receptai\Pavakariai", recipesPav);
            ReadRecipes(@"Receptai\Priešpiečiai", recipesPries);
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
                Color headerColor = Color.FromArgb(146, 208, 80);
                FileInfo savePath = new FileInfo(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Mitybos planai", fileName));
                File.Delete(savePath.FullName);
                using ExcelPackage package = new ExcelPackage(savePath);
                var sheet = package.Workbook.Worksheets.Add("Planas");

                sheet.SelectedRange[1, 1, 100, 5].Style.Font.Size = 10;
                sheet.SelectedRange[1, 1, 100, 5].Style.WrapText = true;
                sheet.SelectedRange[1, 1, 100, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.SelectedRange[1, 1, 100, 5].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                DisplayLogo(sheet);

                sheet.Cells["A2"].Value = "Valgiai";
                sheet.Cells["B2"].Value = "Pirmadienis, Antradienis";
                sheet.Cells["C2"].Value = "Trečiadienis, Ketvirtadienis";
                sheet.Cells["D2"].Value = "Penktadienis, Šeštadienis";
                sheet.Cells["E2"].Value = "Sekmadienis";

                sheet.Column(1).Width = 9;
                sheet.Column(2).Width = 18;
                sheet.Column(3).Width = 18;
                sheet.Column(4).Width = 18;
                sheet.Column(5).Width = 18;

                sheet.Cells[2, 1, 2, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[2, 1, 2, 5].Style.Fill.BackgroundColor.SetColor(headerColor);
                sheet.Cells[3, 1, 8, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[3, 1, 8, 1].Style.Fill.BackgroundColor.SetColor(headerColor);

                sheet.Cells[2, 1, 8, 5].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                sheet.Cells[2, 1, 8, 5].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                sheet.Cells[2, 1, 8, 5].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                sheet.Cells[2, 1, 8, 5].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                DisplayMeal(sheet, package);

                package.Save();

                MessageBox.Show("Dokumentas sukurtas sėkmingai", "Alert", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception e)
            {
                if (e.Source == "System.IO.FileSystem")
                {
                    MessageBox.Show("Uždarykite excel dokumentą. " + e.Message, "Alert", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                else if (e.Message == "Object reference not set to an instance of an object.")
                {
                    MessageBox.Show("Nepasirinkti visi patiekalai. " + e.Message, "Alert", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                else
                {
                    MessageBox.Show("Nežinoma klaida. " + e.Message, "Alert", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void DisplayLogo(ExcelWorksheet sheet)
        {
            FileInfo logo = new FileInfo(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logo.PNG"));
            OfficeOpenXml.Drawing.ExcelPicture pic = sheet.Drawings.AddPicture("logo", logo);
            sheet.SelectedRange[1, 3, 1, 5].Merge = true;
            sheet.Cells[1, 3, 1, 5].Value = "Svarbu!\nŠis mitybos planas yra individualus ir skirtas tik jį užsakiusiam klientui. Dalindamiesi šiuo mitybos planu rizikuojate kito asmens sveikata.";
            sheet.Row(1).Height = 60;
            pic.SetSize(50);
            pic.SetPosition(0, 0);
        }

        private void DisplayMeal(ExcelWorksheet sheet, ExcelPackage package)
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

                GoToNextPage(sheet);
                DisplayRecipes(selectedRecipes, 10, sheet, package);
            }
            catch (Exception)
            {
                throw;
            }
        }

        private void GoToNextPage(ExcelWorksheet sheet)
        {
            double pageHeight = 250;
            double totalHeight = 0;

            for (int row = 1; row <= 8; row++)
            {
                totalHeight += sheet.Row(row).Height;
            }

            if (totalHeight > pageHeight)
                return;

            sheet.Row(9).Height = pageHeight - totalHeight;
        }

        private void DisplayRecipes(List<Recipe> selectedRecipes, int row, ExcelWorksheet sheet, ExcelPackage package)
        {
            //Priešpiečiai
            sheet.Cells[row, 1, row, 5].Merge = true;
            sheet.Cells[row, 1, row, 5].Value = "Priešpiečiai (kas dieną pasirinkite vieną)";
            sheet.Cells[row, 1, row, 5].Style.Border.Top.Style = ExcelBorderStyle.Thick;
            sheet.Cells[row, 1, row, 5].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
            row++;

            foreach (Recipe recipe in recipesPries)
            {
                sheet.Cells[row, 1, row, 5].Merge = true;
                sheet.Cells[row, 1, row, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                sheet.Cells[row, 1, row, 5].RichText.Add(recipe.Title + "\n\n").Bold = true;
                sheet.Cells[row, 1, row, 5].RichText.Add(recipe.Ingredients).Bold = false;
                sheet.Row(row).Height = recipe.IngredientsHeight();
                row++;
            }

            //Pavakariai
            sheet.Cells[row, 1, row, 5].Merge = true;
            sheet.Cells[row, 1, row, 5].Value = "Pavakariai (kas dieną pasirinkite vieną)";
            sheet.Cells[row, 1, row, 5].Style.Border.Top.Style = ExcelBorderStyle.Thick;
            sheet.Cells[row, 1, row, 5].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
            row++;

            foreach (Recipe recipe in recipesPav)
            {
                sheet.Cells[row, 1, row, 5].Merge = true;
                sheet.Cells[row, 1, row, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                sheet.Cells[row, 1, row, 5].RichText.Add(recipe.Title + "\n\n").Bold = true;
                sheet.Cells[row, 1, row, 5].RichText.Add(recipe.Ingredients).Bold = false;
                sheet.Row(row).Height = recipe.IngredientsHeight();
                row++;
            }

            //Visi receptai išvardijami
            sheet.Cells[row, 1, row, 5].Merge = true;
            sheet.Cells[row, 1, row, 5].Value = "Receptai";
            sheet.Cells[row, 1, row, 5].Style.Border.Top.Style = ExcelBorderStyle.Thick;
            sheet.Cells[row, 1, row, 5].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
            row++;

            foreach (Recipe recipe in selectedRecipes)
            {
                if (recipe.Description.Length > 1)
                {
                    sheet.Cells[row, 1, row, 5].Merge = true;
                    sheet.Cells[row, 1, row, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    sheet.Cells[row, 1, row, 5].RichText.Add(recipe.Title + " - ").Bold = true;
                    sheet.Cells[row, 1, row, 5].RichText.Add(recipe.Description).Bold = false;
                    sheet.Row(row).Height = recipe.DescriptionHeight();
                    row++;
                }
            }

            foreach (Recipe recipe in recipesPries)
            {
                if (recipe.Description.Length > 1)
                {
                    sheet.Cells[row, 1, row, 5].Merge = true;
                    sheet.Cells[row, 1, row, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    sheet.Cells[row, 1, row, 5].RichText.Add(recipe.Title + " - ").Bold = true;
                    sheet.Cells[row, 1, row, 5].RichText.Add(recipe.Description).Bold = false;
                    sheet.Row(row).Height = recipe.DescriptionHeight();
                    row++;
                }
            }

            foreach (Recipe recipe in recipesPav)
            {
                if (recipe.Description.Length > 1)
                {
                    sheet.Cells[row, 1, row, 5].Merge = true;
                    sheet.Cells[row, 1, row, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    sheet.Cells[row, 1, row, 5].RichText.Add(recipe.Title + " - ").Bold = true;
                    sheet.Cells[row, 1, row, 5].RichText.Add(recipe.Description).Bold = false;
                    sheet.Row(row).Height = recipe.DescriptionHeight();
                    row++;
                }
            }
        }

        private void AddRecipeIfUnique(List<Recipe> list, Recipe recipe)
        {
            if (!list.Contains(recipe))
                list.Add(recipe);
        }

        private void FillComboBoxes()
        {
            List<Recipe> breakfast = FilterByType("Pusryčiai");
            List<Recipe> lunch = FilterByType("Pietūs");
            List<Recipe> dinner = FilterByType("Vakarienė");

            comboBox.ItemsSource = breakfast;
            comboBox1.ItemsSource = breakfast;
            comboBox2.ItemsSource = breakfast;
            comboBox3.ItemsSource = breakfast;
            comboBox_Copy.ItemsSource = lunch;
            comboBox1_Copy.ItemsSource = lunch;
            comboBox2_Copy.ItemsSource = lunch;
            comboBox3_Copy.ItemsSource = lunch;
            comboBox_Copy1.ItemsSource = dinner;
            comboBox1_Copy1.ItemsSource = dinner;
            comboBox2_Copy1.ItemsSource = dinner;
            comboBox3_Copy1.ItemsSource = dinner;
        }
        private void ReadRecipes(string folderName, List<Recipe> list, string type = "")
        {
            string[] paths = Directory.GetFiles(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, folderName));
            foreach (string path in paths)
            {
                list.Add(new Recipe(path, type));
            }
        }

        private List<Recipe> FilterByType(string type)
        {
            List<Recipe> list = new List<Recipe>();

            foreach (Recipe recipe in recipes)
            {
                if (recipe.Type == type)
                {
                    list.Add(recipe);
                }
            }

            return list;
        }
    }
}
