using System;
using System.Collections.Generic;

namespace MitybosPlanas
{
    class Recipe
    {
        public string Title { get; set; }
        public string Description { get; set; }
        private List<String> ingredients;

        public Recipe(string title, string description, List<string> ingredients)
        {
            Title = title;
            Description = description;
            this.ingredients = ingredients;
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
