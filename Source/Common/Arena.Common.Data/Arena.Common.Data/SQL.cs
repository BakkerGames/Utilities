// SQL.cs - 08/28/2017

using System;
using System.Data.SqlClient;

namespace Arena.Common.Data
{
    public static class SQL
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

        public static SqlConnection GetDataConnection(DataPlug dp)
        {
            SqlConnection dc = new SqlConnection()
            {
                ConnectionString = dp.ConnectionString
            };
            return dc;
        }

        public static string GetConnectionString(string server, string database)
        {
            SqlConnectionStringBuilder result = new SqlConnectionStringBuilder()
            {
                DataSource = server,
                InitialCatalog = database,
                IntegratedSecurity = true,
                PersistSecurityInfo = false,
                Encrypt = false,
                ConnectTimeout = SQLTimeoutSeconds,
                Pooling = true
            };
            return result.ToString();
        }
    }
}
