// SQL.Fix.cs - 06/08/2018

using System.Text;

namespace Arena.Common.Data
{
    public static partial class SQL
    {
        public static string FixSubquery(string value)
        {
            if (value.Contains(" = NULL"))
            {
                return value.Replace(" = NULL", " IS NULL");
            }
            else if (value.Contains(" <> NULL"))
            {
                return value.Replace(" <> NULL", " IS NOT NULL");
            }
            return value;
        }

        public static string AdjustWildcards(string value)
        {
            string result = value;
            if (!string.IsNullOrEmpty(result))
            {
                if (result.Contains("*"))
                {
                    result = result.Replace("*", "%");
                }
                if (result.Contains("?"))
                {
                    result = result.Replace("?", "_");
                }
            }
            return result;
        }

        public static string StringToSQLQuoted_IDRIS(string value)
        {
            StringBuilder result = new StringBuilder();
            if (value != null)
            {
                foreach (char c in value)
                {
                    if (c < 32 || c > 126) // not space to tilde
                    {
                        result.Append(' '); // change to space
                    }
                    else
                    {
                        result.Append(c);
                    }
                }
            }
            return StringToSQLQuoted(result.ToString());
        }
    }
}
