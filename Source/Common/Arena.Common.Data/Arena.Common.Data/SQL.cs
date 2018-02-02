// SQL.cs - 12/18/2017

using System;

namespace Arena.Common.Data
{
    public static partial class SQL
    {
        public static int SQLTimeoutSeconds = 30;

        public static string StringToSQL(string value)
        {
            if (value == null)
            {
                return "NULL";
            }
            else if (value.Contains("'"))
            {
                return value.Replace("'", "''");
            }
            else
            {
                return value;
            }
        }

        public static string StringToSQLQuoted(string value)
        {
            if (value == null)
            {
                return "NULL";
            }
            else if (value.Contains("'"))
            {
                return $"'{value.Replace("'", "''")}'";
            }
            else
            {
                return $"'{value}'";
            }
        }

        public static string GuidToSQLQuoted(Guid value)
        {
            return $"'{value}'";
        }

        public static string GuidToSQLQuoted(Guid? value)
        {
            if (!value.HasValue)
            {
                return "NULL";
            }
            else
            {
                return $"'{value}'";
            }
        }

        public static string DateTimeToSQLQuoted(DateTime value)
        {
            return $"'{value}'";
        }

        public static string DateTimeToSQLQuoted(DateTime? value)
        {
            if (!value.HasValue)
            {
                return "NULL";
            }
            else
            {
                return $"'{value}'";
            }
        }

        public static string BooleanToSQLQuoted(bool value)
        {
            if (value)
            {
                return "1";
            }
            else
            {
                return "0";
            }
        }

        public static string BooleanToSQLQuoted(bool? value)
        {
            if (!value.HasValue)
            {
                return "NULL";
            }
            else if (value.Value)
            {
                return "1";
            }
            else
            {
                return "0";
            }
        }

        public static string NumberToSQL(int? value)
        {
            if (!value.HasValue)
            {
                return "NULL";
            }
            return value.ToString();
        }

        public static string NumberToSQL(long? value)
        {
            if (!value.HasValue)
            {
                return "NULL";
            }
            return value.ToString();
        }

        public static string NumberToSQL(decimal? value)
        {
            if (!value.HasValue)
            {
                return "NULL";
            }
            return value.ToString();
        }
    }
}
