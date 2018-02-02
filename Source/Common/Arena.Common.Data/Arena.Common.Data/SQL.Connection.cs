// SQL.Connection.cs - 12/18/2017

using System.Data.SqlClient;

namespace Arena.Common.Data
{
    public static partial class SQL
    {
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
