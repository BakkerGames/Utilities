// DataPlug.cs - 10/13/2017

using Arena.Common.Errors;
using Arena.Common.Settings;
using System;

namespace Arena.Common.Data
{
    public class DataPlug
    {
        public DataPlug(string productFamily)
        {
            if (string.IsNullOrEmpty(productFamily) )
            {
                throw new SystemException(ErrorHandler.FixMessage("ProductFamily not specified"));
            }
            _productFamily = productFamily;
            _serverName = DataSettings.GetServerName(productFamily);
            _databaseName = DataSettings.GetDatabaseName(productFamily);
            _connectionString = SQL.GetConnectionString(_serverName, _databaseName);
        }

        private string _productFamily;
        public string ProductFamily
        {
            get
            {
                return _productFamily;
            }
        }

        private string _serverName;
        public string ServerName
        {
            get
            {
                return _serverName;
            }
        }

        private string _databaseName;
        public string DatabaseName
        {
            get
            {
                return _databaseName;
            }
        }

        private string _connectionString;
        public string ConnectionString
        {
            get
            {
                return _connectionString;
            }
        }
    }
}
