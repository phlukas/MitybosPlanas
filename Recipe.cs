using System;

namespace MitybosPlanas
{
    class Recipe
    {
        public string Title { get; set; }
        public string Description { get; set; }
        public string Ingredients { get; set; }
        public string Type { get; set; }

        public Recipe(string filePath, string type)
        {
            string ingredients;
            Description = InOut.ReadRecipe(filePath, out ingredients);
            Ingredients = ingredients;
            Type = type;
            SetTitle(filePath);
        }

        private void SetTitle(string filePath)
        {
            int titleStart = filePath.LastIndexOf(@"\");
            filePath = filePath.Remove(0, titleStart + 1);
            int titleEnd = filePath.LastIndexOf(".");
            Title = filePath.Remove(titleEnd);
        }

        public double IngredientsHeight()
        {
            return Ingredients.Split('\n').Length * 15 + 30; //+30 nes pavadinimas + tuscia eilute
        }

        public double DescriptionHeight()
        {
            string[] lines = Description.Split('\n');
            int longLines = 0;

            foreach (string line in lines)
            {
                if (line.Length > 100)
                {
                    longLines += (int)Math.Ceiling((double)line.Length / (double)100);
                }
            }

            return lines.Length * 15 + longLines * 15;
        }
    }
}
