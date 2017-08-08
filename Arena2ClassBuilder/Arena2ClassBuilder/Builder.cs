// Builder.cs - 08/08/2017

using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;

namespace Arena2ClassBuilder
{
    static public class Builder
    {
        public static string DoBuildClass(FileInfo fi, string productFamily)
        {
            string[] lines = File.ReadAllLines(fi.FullName);
            List<FieldItem> fields = new List<FieldItem>();
            bool inFields = false;
            bool afterFields = false;
            bool hasIDCode = false;
            foreach (string currLine in lines)
            {
                if (afterFields)
                {
                    continue;
                }
                string currLineUpper = currLine.ToUpper().Trim();
                if (!inFields && !afterFields)
                {
                    if (currLineUpper.Contains("CREATE TABLE"))
                    {
                        inFields = true;
                    }
                    continue;
                }
                if (inFields && !afterFields)
                {
                    if (currLineUpper.Contains(") ON PRIMARY"))
                    {
                        afterFields = true;
                        continue;
                    }
                }
                if (!currLineUpper.StartsWith("[") || !currLineUpper.Contains("NULL"))
                {
                    continue;
                }
                string[] tokens = currLine.Trim().Replace("  ", " ").Split(' ');
                bool firstToken = true;
                bool secondToken = false;
                FieldItem currFieldItem = new FieldItem();
                string tempToken;
                string appendFromLast = "";
                foreach (string currToken in tokens)
                {
                    tempToken = currToken;
                    if (!string.IsNullOrEmpty(appendFromLast))
                    {
                        tempToken = appendFromLast + tempToken;
                        appendFromLast = "";
                    }
                    if (tempToken.Contains("[") && !tempToken.Contains("]"))
                    {
                        appendFromLast = tempToken + "_";
                        continue;
                    }
                    if (firstToken)
                    {
                        // fieldname
                        tempToken = tempToken.Substring(1, tempToken.Length - 2); // remove []
                        currFieldItem.FieldName = tempToken;
                        if (tempToken.Equals("IDCode", StringComparison.OrdinalIgnoreCase))
                        {
                            hasIDCode = true;
                        }
                        firstToken = false;
                        secondToken = true;
                        continue;
                    }
                    tempToken = tempToken.ToUpper();
                    tempToken = tempToken.Replace("[", "").Replace("]", "");
                    tempToken = tempToken.Replace("(", " ").Replace(")", "");
                    tempToken = tempToken.Replace(",", "");
                    if (string.Equals(tempToken, "NOT", StringComparison.OrdinalIgnoreCase))
                    {
                        currFieldItem.NotNull = true;
                    }
                    if (secondToken)
                    {
                        string[] fieldType = tempToken.Split(' ');
                        currFieldItem.FieldType = fieldType[0];
                        if (fieldType.GetUpperBound(0) > 0)
                        {
                            currFieldItem.FieldLen = fieldType[1];
                        }
                        else
                        {
                            currFieldItem.FieldLen = null;
                        }
                        secondToken = false;
                    }
                }
                if (productFamily.Equals("Advantage", StringComparison.OrdinalIgnoreCase) 
                    || !IgnoreField(currFieldItem))
                {
                    fields.Add(currFieldItem);
                }
            }

            // build the class from the known information
            string result = "";
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resource = asm.GetManifestResourceNames();
            string streamName;
            if (productFamily.Equals("IDRIS", StringComparison.OrdinalIgnoreCase))
            {
                streamName = "Arena2ClassBuilder.Resources.BlankIDRIS2DataClass.txt";
            }
            else if (productFamily.Equals("Advantage", StringComparison.OrdinalIgnoreCase))
            {
                streamName = "Arena2ClassBuilder.Resources.BlankAdvantage2DataClass.txt";
            }
            else if (!hasIDCode)
            {
                streamName = "Arena2ClassBuilder.Resources.BlankArena2DataClassNoIDCode.txt";
            }
            else
            {
                streamName = "Arena2ClassBuilder.Resources.BlankArena2DataClass.txt";
            }
            Stream blankTemplateStream = asm.GetManifestResourceStream(streamName);
            StreamReader sr = new StreamReader(blankTemplateStream);
            result = sr.ReadToEnd();
            sr.Close();

            // find names for replacing below
            string schemaNameSQL = fi.Name.Substring(0, fi.Name.IndexOf("."));
            string schemaName = schemaNameSQL;
            if (schemaName.Equals("dbo"))
            {
                schemaName = productFamily;
            }
            string tableName = fi.Name.Substring(schemaNameSQL.Length + 1, fi.Name.Length - schemaNameSQL.Length - 11);
            string className = $"{schemaName}_{tableName}_DataAccess";

            // replace all special tokens in template with field info
            result = result.Replace("$SCHEMANAMESQL$", schemaNameSQL);
            result = result.Replace("$SCHEMANAME$", schemaName);
            result = result.Replace("$TABLENAME$", tableName);
            result = result.Replace("$CLASSNAME$", className);
            result = result.Replace("$ORDDEFS$\r\n", GetOrdinalDefs(fields));
            result = result.Replace("$PROPERTIES$\r\n", GetProperties(fields));
            result = result.Replace("$GETFIELDLIST$\r\n", GetFieldList(fields));
            result = result.Replace("$TOSTRINGFIELDS$\r\n", GetToStringFields(fields));
            result = result.Replace("$GETINSERTFIELDLIST$\r\n", GetInsertFieldList(fields));
            result = result.Replace("$GETINSERTVALUELIST$\r\n", GetInsertValueList(fields));
            result = result.Replace("$GETUPDATEVALUELIST$\r\n", GetUpdateValueList(fields));
            result = result.Replace("$SETORDINALS$\r\n", GetSetOrdinals(fields, productFamily, hasIDCode));
            result = result.Replace("$FILLFIELDS$\r\n", GetFillFields(fields));

            return result;
        }

        private static string GetFillFields(List<FieldItem> fields)
        {
            StringBuilder result = new StringBuilder();
            foreach (FieldItem currFieldItem in fields)
            {
                result.Append("            obj.");
                result.Append(currFieldItem.FieldName);
                result.Append(" = ");
                if (!currFieldItem.NotNull)
                {
                    result.Append("dr.IsDBNull(_ord");
                    result.Append(currFieldItem.FieldName);
                    result.Append(") ? null : ");
                    switch (currFieldItem.FieldType)
                    {
                        case "BIT":
                            result.Append("(bool?)");
                            break;
                        case "INT":
                        case "SMALLINT":
                        case "TINYINT":
                            result.Append("(int?)");
                            break;
                        case "DECIMAL":
                        case "MONEY":
                        case "SMALLMONEY":
                        case "FLOAT":
                        case "REAL":
                            result.Append("(decimal?)");
                            break;
                        case "DATE":
                        case "DATETIME":
                        case "DATETIME2":
                            result.Append("(DateTime?)");
                            break;
                        case "CHAR":
                        case "NCHAR":
                        case "VARCHAR":
                        case "NVARCHAR":
                            // needs no conversion
                            break;
                        case "TIMESTAMP":
                            result.Append("(long?)");
                            break;
                        case "UNIQUEIDENTIFIER":
                            result.Append("(Guid?)");
                            break;
                        default:
                            result.Append("###");
                            break;
                    }
                }
                switch (currFieldItem.FieldType)
                {
                    case "MONEY":
                    case "SMALLMONEY":
                        result.Append("(decimal)");
                        break;
                    case "REAL":
                        result.Append("(decimal)");
                        break;
                }
                result.Append("dr.");
                switch (currFieldItem.FieldType)
                {
                    case "BIT":
                        result.Append("GetBoolean");
                        break;
                    case "TINYINT":
                        result.Append("GetByte");
                        break;
                    case "SMALLINT":
                        result.Append("GetInt16");
                        break;
                    case "INT":
                        result.Append("GetInt32");
                        break;
                    case "LONG":
                        result.Append("GetInt64");
                        break;
                    case "DECIMAL":
                        result.Append("GetDecimal");
                        break;
                    case "MONEY":
                    case "SMALLMONEY":
                        result.Append("GetSqlMoney");
                        break;
                    case "FLOAT":
                        result.Append("GetDouble");
                        break;
                    case "REAL":
                        result.Append("GetFloat");
                        break;
                    case "DATE":
                    case "DATETIME":
                    case "DATETIME2":
                        result.Append("GetDateTime");
                        break;
                    case "CHAR":
                    case "NCHAR":
                    case "VARCHAR":
                    case "NVARCHAR":
                        result.Append("GetString");
                        break;
                    case "TIMESTAMP":
                        result.Append("GetInt64");
                        break;
                    case "UNIQUEIDENTIFIER":
                        result.Append("GetGuid");
                        break;
                    default:
                        result.Append("###");
                        break;
                }
                result.Append("(_ord");
                result.Append(currFieldItem.FieldName);
                result.AppendLine(");");
            }
            return result.ToString();
        }

        private static string GetUpdateValueList(List<FieldItem> fields)
        {
            StringBuilder result = new StringBuilder();
            bool firstField = true;
            foreach (FieldItem currFieldItem in fields)
            {
                result.Append("            sb.Append(\"");
                if (!firstField)
                {
                    result.Append(", ");
                }
                result.Append("[");
                result.Append(currFieldItem.FieldName);
                result.AppendLine("] = \");");
                result.Append("            sb.Append(");
                switch (currFieldItem.FieldType)
                {
                    case "BIT":
                        result.Append("SQL.BooleanToSQLQuoted(obj.");
                        result.Append(currFieldItem.FieldName);
                        result.Append(")");
                        break;
                    case "CHAR":
                    case "NCHAR":
                    case "VARCHAR":
                    case "NVARCHAR":
                        result.Append("SQL.StringToSQLQuoted(obj.");
                        result.Append(currFieldItem.FieldName);
                        result.Append(")");
                        break;
                    case "DATE":
                    case "DATETIME":
                    case "DATETIME2":
                        result.Append("SQL.DateTimeToSQLQuoted(obj.");
                        result.Append(currFieldItem.FieldName);
                        result.Append(")");
                        break;
                    default:
                        result.Append("obj.");
                        result.Append(currFieldItem.FieldName);
                        if (currFieldItem.NotNull)
                        {
                            result.Append(".ToString()");
                        }
                        else
                        {
                            result.Append("?.ToString() ?? \"NULL\"");
                        }
                        break;
                }
                result.AppendLine(");");
                firstField = false;
            }
            return result.ToString();
        }

        private static string GetInsertValueList(List<FieldItem> fields)
        {
            StringBuilder result = new StringBuilder();
            bool firstField = true;
            foreach (FieldItem currFieldItem in fields)
            {
                if (!firstField)
                {
                    result.AppendLine("            sb.Append(\", \");");
                }
                result.Append("            sb.Append(");
                switch (currFieldItem.FieldType)
                {
                    case "BIT":
                        result.Append("SQL.BooleanToSQLQuoted(obj.");
                        result.Append(currFieldItem.FieldName);
                        result.Append(")");
                        break;
                    case "CHAR":
                    case "NCHAR":
                    case "VARCHAR":
                    case "NVARCHAR":
                        result.Append("SQL.StringToSQLQuoted(obj.");
                        result.Append(currFieldItem.FieldName);
                        result.Append(")");
                        break;
                    case "DATE":
                    case "DATETIME":
                    case "DATETIME2":
                        result.Append("SQL.DateTimeToSQLQuoted(obj.");
                        result.Append(currFieldItem.FieldName);
                        result.Append(")");
                        break;
                    default:
                        result.Append("obj.");
                        result.Append(currFieldItem.FieldName);
                        if (currFieldItem.NotNull)
                        {
                            result.Append(".ToString()");
                        }
                        else
                        {
                            result.Append("?.ToString() ?? \"NULL\"");
                        }
                        break;
                }
                result.AppendLine(");");
                firstField = false;
            }
            return result.ToString();
        }

        private static string GetInsertFieldList(List<FieldItem> fields)
        {
            StringBuilder result = new StringBuilder();
            bool firstField = true;
            foreach (FieldItem currFieldItem in fields)
            {
                result.Append("            sb.Append(\"");
                if (!firstField)
                {
                    result.Append(", ");
                }
                result.Append("[");
                result.Append(currFieldItem.FieldName);
                result.AppendLine("]\");");
                firstField = false;
            }
            return result.ToString();
        }

        private static string GetFieldList(List<FieldItem> fields)
        {
            StringBuilder result = new StringBuilder();
            foreach (FieldItem currFieldItem in fields)
            {
                if (currFieldItem.FieldName.Equals("RowVersion", StringComparison.OrdinalIgnoreCase) ||
                    currFieldItem.FieldName.Equals("Timestamp", StringComparison.OrdinalIgnoreCase))
                {
                    result.Append("            sb.Append(\", CONVERT(BIGINT,[");
                    result.Append(currFieldItem.FieldName);
                    result.Append("]) AS [");
                    result.Append(currFieldItem.FieldName);
                    result.AppendLine("]\");");
                }
                else
                {
                    result.Append("            sb.Append(\", [");
                    result.Append(currFieldItem.FieldName);
                    result.AppendLine("]\");");
                }
            }
            return result.ToString();
        }

        private static string GetToStringFields(List<FieldItem> fields)
        {
            StringBuilder result = new StringBuilder();
            foreach (FieldItem currFieldItem in fields)
            {
                result.Append("                { \"");
                result.Append(currFieldItem.FieldName);
                result.Append("\", ");
                result.Append(currFieldItem.FieldName);
                result.AppendLine(" },");
            }
            return result.ToString();
        }

        private static string GetSetOrdinals(List<FieldItem> fields, string productFamily, bool hasIDCode)
        {
            StringBuilder result = new StringBuilder();
            int nextOrdinal;
            if (productFamily.Equals("IDRIS", StringComparison.OrdinalIgnoreCase))
            {
                nextOrdinal = 4; // IDRIS only has 4 header fields
            }
            else if (productFamily.Equals("Advantage", StringComparison.OrdinalIgnoreCase))
            {
                nextOrdinal = 0; // Advantage has no header fields
            }
            else if (!hasIDCode)
            {
                nextOrdinal = 4; // Arena without IDCode only has 4 header fields
            }
            else
            {
                nextOrdinal = 5; // Arena2 has 5 header fields
            }
            foreach (FieldItem currFieldItem in fields)
            {
                result.Append("            _ord");
                result.Append(currFieldItem.FieldName);
                result.Append(" = ");
                result.Append(nextOrdinal++);
                result.AppendLine(";");
            }
            return result.ToString();
        }

        private static string GetProperties(List<FieldItem> fields)
        {
            StringBuilder result = new StringBuilder();
            foreach (FieldItem currFieldItem in fields)
            {
                result.Append("        public ");
                switch (currFieldItem.FieldType)
                {
                    case "BIT":
                        result.Append("bool");
                        if (!currFieldItem.NotNull)
                        {
                            result.Append("?");
                        }
                        break;
                    case "INT":
                    case "SMALLINT":
                    case "TINYINT":
                        result.Append("int");
                        if (!currFieldItem.NotNull)
                        {
                            result.Append("?");
                        }
                        break;
                    case "DECIMAL":
                    case "MONEY":
                    case "SMALLMONEY":
                    case "FLOAT":
                    case "REAL":
                        result.Append("decimal");
                        if (!currFieldItem.NotNull)
                        {
                            result.Append("?");
                        }
                        break;
                    case "DATE":
                    case "DATETIME":
                    case "DATETIME2":
                        result.Append("DateTime");
                        if (!currFieldItem.NotNull)
                        {
                            result.Append("?");
                        }
                        break;
                    case "CHAR":
                    case "NCHAR":
                    case "VARCHAR":
                    case "NVARCHAR":
                        result.Append("string");
                        break;
                    case "TIMESTAMP":
                        result.Append("long");
                        if (!currFieldItem.NotNull)
                        {
                            result.Append("?");
                        }
                        break;
                    case "UNIQUEIDENTIFIER":
                        result.Append("Guid");
                        if (!currFieldItem.NotNull)
                        {
                            result.Append("?");
                        }
                        break;
                    default:
                        result.Append("###");
                        break;
                }
                result.Append(" ");
                result.Append(currFieldItem.FieldName);
                result.AppendLine(" { get; set; }");
            }
            return result.ToString();
        }

        private static string GetOrdinalDefs(List<FieldItem> fields)
        {
            StringBuilder result = new StringBuilder();
            foreach (FieldItem currFieldItem in fields)
            {
                result.Append("        private static int _ord");
                result.Append(currFieldItem.FieldName);
                result.AppendLine(";");
            }
            return result.ToString();
        }

        private static bool IgnoreField(FieldItem currFieldItem)
        {
            // these fields are standard on every table
            if (string.Equals(currFieldItem.FieldName, "ID", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
            if (string.Equals(currFieldItem.FieldName, "IDCode", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
            if (string.Equals(currFieldItem.FieldName, "RowVersion", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(currFieldItem.FieldName, "TimeStamp", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
            if (string.Equals(currFieldItem.FieldName, "LastChanged", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
            if (string.Equals(currFieldItem.FieldName, "ChangedBy", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
            // IDRIS standard fields
            if (string.Equals(currFieldItem.FieldName, "REC", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
            if (string.Equals(currFieldItem.FieldName, "PACKED_DATA", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
            return false;
        }
    }
}
