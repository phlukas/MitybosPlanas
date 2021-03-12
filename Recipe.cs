using System;
using System.Collections.Generic;

namespace MitybosPlanas
{
    class Recipe
    {
        public string Title { get; set; }
        public string Description { get; set; }
        private List<String> ingredients = new List<string>();

        public Recipe(string filePath)
        {
            Description = InOut.ReadRecipe(filePath, ingredients);
            SetTitle(filePath);
        }

        private void SetTitle(string filePath)
        {
            int titleStart = filePath.LastIndexOf(@"\");
            filePath = filePath.Remove(0, titleStart + 1);
            int titleEnd = filePath.LastIndexOf(".");
            Title = filePath.Remove(titleEnd);
        }

        public String GetIngredient(int i)
        {
            return ingredients[i];
        }

        public int IngredientsCount()
        {
            return ingredients.Count;
        }
    }
}
