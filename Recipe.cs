namespace MitybosPlanas
{
    class Recipe
    {
        public string Title { get; set; }
        public string Description { get; set; }
        public string Ingredients { get; set; }

        public Recipe(string filePath)
        {
            string ingredients;
            Description = InOut.ReadRecipe(filePath, out ingredients);
            Ingredients = ingredients;
            SetTitle(filePath);
        }

        private void SetTitle(string filePath)
        {
            int titleStart = filePath.LastIndexOf(@"\");
            filePath = filePath.Remove(0, titleStart + 1);
            int titleEnd = filePath.LastIndexOf(".");
            Title = filePath.Remove(titleEnd);
        }
    }
}
