using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace MitybosPlanas
{
    static class InOut
    {
        public static string ReadRecipe(string path, out string ingredients)
        {
            StringBuilder sb = new StringBuilder();
            StringBuilder sb2 = new StringBuilder();

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
                    sb2.Append(line);
                    sb2.Append("\n");
                    line = sr.ReadLine();
                }
                ingredients = sb2.ToString();
            }

            return sb.ToString();
        }
    }
}
