// SQL.Fix.cs - 01/23/2018

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
    }
}
