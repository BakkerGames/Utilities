// Builder.cs - 08/01/2018

using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;

namespace Arena2ClassBuilder
{
    static public class Builder
    {
        internal static List<string> ignoreFieldList = new List<string>();
        internal static string baseClassName;

        public static string DoBuildClass(FileInfo fi, string productFamily)
        {
            // check for info file
            ignoreFieldList.Clear();
            baseClassName = "";
            string infoFilename = fi.FullName.Substring(0, fi.FullName.Length - 4) + ".info";
            if (File.Exists(infoFilename))
            {
                foreach (string tempLine in File.ReadAllLines(infoFilename))
                {
                    if (tempLine.StartsWith("#BASECLASS#"))
                    {
                        baseClassName = tempLine.Substring(11).Trim();
                    }
                    else if (tempLine.StartsWith("#IGNORE#"))
                    {
                        ignoreFieldList.Add(tempLine.Substring(8).Trim());
                    }
                }
            }
            string[] lines = File.ReadAllLines(fi.FullName);
            List<FieldItem> fields = new List<FieldItem>();
            List<FieldItem> fieldsJson = new List<FieldItem>();
            bool inFields = false;
            bool afterFields = false;
            bool hasIdCode = false;
            string identityFieldname = "";
            foreach (string currLine in lines)
            {
                string currLineUpper = currLine.ToUpper().Trim();
                if (afterFields)
                {
                    // look for "DEFAULT expression FOR field"
                    if (currLineUpper.Contains(" DEFAULT ") && currLineUpper.Contains(" FOR "))
                    {
                        int defPos = currLineUpper.IndexOf(" DEFAULT ") + 9;
                        int forPos = currLineUpper.IndexOf(" FOR ") + 5;
                        // get defValue from currLine so the case isn't changed
                        string defValue = currLine.Trim().Substring(defPos, forPos - defPos - 5).Trim();
                        string forValue = currLineUpper.Substring(forPos).Trim();
                        if (forValue.EndsWith(";"))
                        {
                            forValue = forValue.Substring(0, forValue.Length - 1).Trim();
                        }
                        if (forValue.StartsWith("[") && forValue.EndsWith("]"))
                        {
                            forValue = forValue.Substring(1, forValue.Length - 2).Trim();
                        }
                        foreach (FieldItem f in fields)
                        {
                            if (f.FieldName.Equals(forValue, StringComparison.OrdinalIgnoreCase))
                            {
                                f.DefaultValue = defValue;
                                break;
                            }
                        }
                    }
                    continue;
                }
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
                    if (currLineUpper.Trim().StartsWith(") ON PRIMARY")
                        || currLineUpper.Trim().StartsWith(") ON [PRIMARY]"))
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
                        if (tempToken.StartsWith("[") && tempToken.EndsWith("]"))
                        {
                            tempToken = tempToken.Substring(1, tempToken.Length - 2); // remove []
                        }
                        currFieldItem.SQLFieldName = tempToken;
                        if (tempToken[0] >= '0' && tempToken[0] <= '9')
                        {
                            tempToken = $"_{tempToken}";
                        }
                        currFieldItem.FieldName = tempToken;
                        if (tempToken.Equals("IdCode", StringComparison.OrdinalIgnoreCase))
                        {
                            hasIdCode = true;
                        }
                        if (currLineUpper.Contains("IDENTITY"))
                        {
                            identityFieldname = tempToken;
                            currFieldItem.IsIdentity = true;
                        }
                        if (currLineUpper.Contains("TIMESTAMP") || currLineUpper.Contains("ROWVERSION"))
                        {
                            currFieldItem.IsTimestamp = true;
                        }
                        firstToken = false;
                        secondToken = true;
                        continue;
                    }
                    tempToken = tempToken.ToUpper();
                    tempToken = tempToken.Replace("[", "").Replace("]", "");
                    tempToken = tempToken.Replace("(", " ").Replace(")", "");
                    tempToken = tempToken.Replace(",", "");
                    if (string.Equals(tempToken, "NOT", StringComparison.OrdinalIgnoreCase) &&
                        !currLineUpper.Contains("IDENTITY"))
                    {
                        // all identity fields should be nullable in data access class
                        currFieldItem.NotNull = true;
                    }
                    if (secondToken)
                    {
                        string[] fieldType = tempToken.Split(' ');
                        if (currFieldItem.FieldName.Equals("PACKED_DATA", StringComparison.OrdinalIgnoreCase))
                        {
                            currFieldItem.FieldType = "VARCHAR";
                            currFieldItem.FieldLen = "512";
                        }
                        else
                        {
                            currFieldItem.FieldType = fieldType[0];
                            if (fieldType.GetUpperBound(0) > 0)
                            {
                                currFieldItem.FieldLen = fieldType[1];
                            }
                            else
                            {
                                currFieldItem.FieldLen = null;
                            }
                        }
                        secondToken = false;
                    }
                }
                if (!IgnoreField(currFieldItem, fi.Name)
                    || productFamily.Equals("Advantage", StringComparison.OrdinalIgnoreCase)
                    || productFamily.Equals("IDRIS Advantage", StringComparison.OrdinalIgnoreCase))
                {
                    fields.Add(currFieldItem);
                }
                // almost all fields for Json object
                if (!IgnoreFieldJson(currFieldItem, fi.Name))
                {
                    // fix standard field name case issues
                    if (currFieldItem.FieldName.Equals("LASTCHANGED"))
                    {
                        currFieldItem.FieldName = "LastChanged";
                    }
                    if (currFieldItem.FieldName.Equals("CHANGEDBY"))
                    {
                        currFieldItem.FieldName = "ChangedBy";
                    }
                    if (currFieldItem.FieldName.Equals("ROWVERSION"))
                    {
                        currFieldItem.FieldName = "RowVersion";
                    }
                    // add to json field list
                    fieldsJson.Add(currFieldItem);
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
            else if (productFamily.Equals("Advantage", StringComparison.OrdinalIgnoreCase)
                || productFamily.Equals("IDRIS Advantage", StringComparison.OrdinalIgnoreCase))
            {
                if (string.IsNullOrEmpty(identityFieldname))
                {
                    streamName = "Arena2ClassBuilder.Resources.BlankAdvantage2DataClassNoIDNum.txt";
                }
                else
                {
                    streamName = "Arena2ClassBuilder.Resources.BlankAdvantage2DataClass.txt";
                }
            }
            else if (!hasIdCode)
            {
                streamName = "Arena2ClassBuilder.Resources.BlankArena2DataClassNoIdCode.txt";
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
            string tableNameSQL = tableName;
            if (productFamily == "IDRIS Advantage" && tableNameSQL.StartsWith("IDRIS_")
                && tableNameSQL != "IDRIS_EXTRACT_RUNDATE")
            {
                tableNameSQL = tableNameSQL.Substring(6); // remove "IDRIS_"
            }
            if (tableNameSQL.StartsWith("_"))
            {
                tableNameSQL = $"[%{tableNameSQL.Substring(1)}]";
            }
            else if (tableNameSQL.Contains(" "))
            {
                tableNameSQL = $"[{tableNameSQL}]";
            }

            string className = $"{schemaName}_{tableName}_DataAccess";
            if (productFamily == "IDRIS Advantage")
            {
                className = $"{tableName}_DataAccess";
            }

            // replace all special tokens in template with field info
            if (!string.IsNullOrEmpty(baseClassName))
            {
                result = result.Replace("$BASECLASS$", baseClassName);
            }
            else
            {
                result = result.Replace(" : $BASECLASS$", "");
            }
            result = result.Replace("$SCHEMANAMESQL$", schemaNameSQL);
            result = result.Replace("$SCHEMANAME$", schemaName);
            result = result.Replace("$TABLENAMESQL$", tableNameSQL);
            result = result.Replace("$TABLENAME$", tableName);
            result = result.Replace("$CLASSNAME$", className);
            result = result.Replace("$ORDDEFS$\r\n", GetOrdinalDefs(fields));
            result = result.Replace("$PROPERTIES$\r\n", GetProperties(fields));
            result = result.Replace("$GETFIELDLIST$\r\n", GetFieldList(fields,
                (productFamily.Equals("Advantage", StringComparison.OrdinalIgnoreCase)
                || productFamily.Equals("IDRIS Advantage", StringComparison.OrdinalIgnoreCase))));
            result = result.Replace("$TOSTRINGFIELDS$\r\n", GetToStringFields(fieldsJson));
            result = result.Replace("$GETINSERTFIELDLIST$\r\n", GetInsertFieldList(fields));
            if (fi.Name.StartsWith("dbo.xt")) // no special handling
            {
                result = result.Replace("$GETINSERTVALUELIST$\r\n", GetInsertValueList(fields, null));
                result = result.Replace("$GETUPDATEVALUELIST$\r\n", GetUpdateValueList(fields, null));
            }
            else
            {
                result = result.Replace("$GETINSERTVALUELIST$\r\n", GetInsertValueList(fields, productFamily));
                result = result.Replace("$GETUPDATEVALUELIST$\r\n", GetUpdateValueList(fields, productFamily));
            }
            result = result.Replace("$SETORDINALS$\r\n", GetSetOrdinals(fields, productFamily, hasIdCode));
            result = result.Replace("$FILLFIELDS$\r\n", GetFillFields(fields));
            result = result.Replace("$IDENTITY$", identityFieldname);

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
                else
                {
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

        private static string GetUpdateValueList(List<FieldItem> fields, string productFamily)
        {
            StringBuilder result = new StringBuilder();
            bool firstField = true;
            foreach (FieldItem currFieldItem in fields)
            {
                if (currFieldItem.IsIdentity || currFieldItem.IsTimestamp)
                {
                    continue;
                }
                if (currFieldItem.FieldName.Equals("PACKED_DATA", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }
                result.Append("            sb.Append(\"");
                if (!firstField)
                {
                    result.Append(", ");
                }
                result.Append("[");
                result.Append(currFieldItem.SQLFieldName);
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
                        if (productFamily == "IDRIS")
                        {
                            result.Append("SQL.StringToSQLQuoted_IDRIS(obj.");
                        }
                        else
                        {
                            result.Append("SQL.StringToSQLQuoted(obj.");
                        }
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
                    case "UNIQUEIDENTIFIER":
                        result.Append("SQL.GuidToSQLQuoted(obj.");
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

        private static string GetInsertValueList(List<FieldItem> fields, string productFamily)
        {
            StringBuilder result = new StringBuilder();
            bool firstField = true;
            foreach (FieldItem currFieldItem in fields)
            {
                if (currFieldItem.IsIdentity || currFieldItem.IsTimestamp)
                {
                    continue;
                }
                if (currFieldItem.FieldName.Equals("PACKED_DATA", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }
                if (!firstField)
                {
                    result.AppendLine("            sb.Append(\", \");");
                }
                if (!string.IsNullOrEmpty(currFieldItem.DefaultValue))
                {
                    result.Append("            if (obj.");
                    result.Append(currFieldItem.FieldName);
                    result.AppendLine(" == null)");
                    result.AppendLine("            {");
                    result.Append("                sb.Append(\"");
                    result.Append(currFieldItem.DefaultValue);
                    result.AppendLine("\");");
                    result.AppendLine("            }");
                    result.AppendLine("            else");
                    result.AppendLine("            {");
                    result.Append("    ");
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
                        if (productFamily == "IDRIS")
                        {
                            result.Append("SQL.StringToSQLQuoted_IDRIS(obj.");
                        }
                        else
                        {
                            result.Append("SQL.StringToSQLQuoted(obj.");
                        }
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
                    case "UNIQUEIDENTIFIER":
                        result.Append("SQL.GuidToSQLQuoted(obj.");
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
                if (!string.IsNullOrEmpty(currFieldItem.DefaultValue))
                {
                    result.AppendLine("            }");
                }
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
                if (currFieldItem.IsIdentity || currFieldItem.IsTimestamp)
                {
                    continue;
                }
                if (currFieldItem.FieldName.Equals("PACKED_DATA", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }
                result.Append("            sb.Append(\"");
                if (!firstField)
                {
                    result.Append(", ");
                }
                result.Append("[");
                result.Append(currFieldItem.SQLFieldName);
                result.AppendLine("]\");");
                firstField = false;
            }
            return result.ToString();
        }

        private static string GetFieldList(List<FieldItem> fields, bool isAdvantage)
        {
            StringBuilder result = new StringBuilder();
            bool firstField = false;
            if (isAdvantage)
            {
                firstField = true;
            }
            foreach (FieldItem currFieldItem in fields)
            {
                if (currFieldItem.FieldName.Equals("RowVersion", StringComparison.OrdinalIgnoreCase) ||
                    currFieldItem.FieldName.Equals("Timestamp", StringComparison.OrdinalIgnoreCase))
                {
                    result.Append("            sb.Append(\", CONVERT(BIGINT,[");
                    result.Append(currFieldItem.SQLFieldName);
                    result.Append("]) AS [");
                    result.Append(currFieldItem.SQLFieldName);
                    result.AppendLine("]\");");
                }
                else if (currFieldItem.FieldName.Equals("PACKED_DATA", StringComparison.OrdinalIgnoreCase))
                {
                    result.Append("            sb.Append(\", CONVERT(VARCHAR(MAX),[");
                    result.Append(currFieldItem.SQLFieldName);
                    result.Append("],2) AS [");
                    result.Append(currFieldItem.SQLFieldName);
                    result.AppendLine("]\");");
                }
                else
                {
                    result.Append("            sb.Append(\"");
                    if (!firstField)
                    {
                        result.Append(", ");
                    }
                    result.Append("[");
                    result.Append(currFieldItem.SQLFieldName);
                    result.AppendLine("]\");");
                    firstField = false;
                }
            }
            return result.ToString();
        }

        private static string GetToStringFields(List<FieldItem> fields)
        {
            StringBuilder result = new StringBuilder();
            foreach (FieldItem currFieldItem in fields)
            {
                if (result.Length > 0)
                {
                    result.AppendLine(",");
                }
                result.Append("                { \"");
                result.Append(currFieldItem.FieldName);
                result.Append("\", ");
                result.Append(currFieldItem.FieldName);
                result.Append(" }");
            }
            result.AppendLine();
            return result.ToString();
        }

        private static string GetSetOrdinals(List<FieldItem> fields,
                                             string productFamily,
                                             bool hasIdCode)
        {
            StringBuilder result = new StringBuilder();
            int nextOrdinal;
            if (productFamily.Equals("IDRIS", StringComparison.OrdinalIgnoreCase))
            {
                nextOrdinal = 4; // IDRIS only has 4 header fields
            }
            else if (productFamily.Equals("Advantage", StringComparison.OrdinalIgnoreCase)
                || productFamily.Equals("IDRIS Advantage", StringComparison.OrdinalIgnoreCase))
            {
                nextOrdinal = 0; // no consistant header fields
            }
            else if (!hasIdCode)
            {
                nextOrdinal = 4; // Arena without IdCode only has 4 header fields
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

        private static bool IgnoreField(FieldItem currFieldItem, string tableName)
        {
            // Arena standard fields
            if (string.Equals(currFieldItem.FieldName, "ID", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
            if (string.Equals(currFieldItem.FieldName, "IdCode", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
            if (string.Equals(currFieldItem.FieldName, "RowVersion", StringComparison.OrdinalIgnoreCase))
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
            // Advantage standard fields
            if (string.Equals(currFieldItem.FieldName, "IDNum", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
            if (string.Equals(currFieldItem.FieldName, "Timestamp", StringComparison.OrdinalIgnoreCase))
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
                if (string.IsNullOrEmpty(tableName) || tableName != "dbo._SCF.Table.sql")
                {
                    return true;
                }
            }
            // check ignore fields from INFO file
            foreach (string field in ignoreFieldList)
            {
                if (string.Equals(currFieldItem.FieldName, field, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
            return false;
        }

        private static bool IgnoreFieldJson(FieldItem currFieldItem, string tableName)
        {
            // IDRIS standard fields
            if (string.Equals(currFieldItem.FieldName, "PACKED_DATA", StringComparison.OrdinalIgnoreCase))
            {
                if (string.IsNullOrEmpty(tableName) || tableName != "dbo._SCF.Table.sql")
                {
                    return true;
                }
            }
            // check ignore fields from INFO file
            foreach (string field in ignoreFieldList)
            {
                if (string.Equals(currFieldItem.FieldName, field, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
            return false;
        }
    }
}
