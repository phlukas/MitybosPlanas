using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace MitybosPlanas
{
    static class InOut
    {
        public static string ReadRecipe(string path, List<string> ingredients)
        {
            StringBuilder sb = new StringBuilder();

            using (StreamReader sr = new StreamReader(path))
            {
                string line = sr.ReadLine();
                while (line != null && line != "@")
                {
                    sb.Append(line);
                    sb.Append("\n");
                    line = sr.ReadLine();
                }
                while (line != null)
                {
                    ingredients.Add(line);
                    line = sr.ReadLine();
                }
            }

            return sb.ToString();
        }
    }
}
