﻿// Builder.cs - 05/10/2017

using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;

namespace Arena2ClassBuilder
{
    static public class Builder
    {
        public static string DoBuildClass(FileInfo fi, bool isIDRIS)
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
                foreach (string currToken in tokens)
                {
                    tempToken = currToken;
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
                if (!IgnoreField(currFieldItem))
                {
                    fields.Add(currFieldItem);
                }
            }

            // build the class from the known information
            string result = "";
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resource = asm.GetManifestResourceNames();
            string streamName;
            if (isIDRIS)
            {
                streamName = "Arena2ClassBuilder.Resources.BlankIDRIS2DataClass.txt";
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

            // replace all special tokens in template with field info
            string schemaName = fi.Name.Substring(0, fi.Name.IndexOf("."));
            string tableName = fi.Name.Substring(schemaName.Length + 1, fi.Name.Length - schemaName.Length - 11);
            string className = $"{schemaName}_{tableName}";
            result = result.Replace("$SCHEMANAME$", schemaName);
            result = result.Replace("$TABLENAME$", tableName);
            result = result.Replace("$CLASSNAME$", className);
            result = result.Replace("$ORDDEFS$\r\n", GetOrdinalDefs(fields));
            result = result.Replace("$PROPERTIES$\r\n", GetProperties(fields));
            result = result.Replace("$GETFIELDLIST$\r\n", GetFieldList(fields));
            result = result.Replace("$GETINSERTFIELDLIST$\r\n", GetInsertFieldList(fields));
            result = result.Replace("$GETINSERTVALUELIST$\r\n", GetInsertValueList(fields));
            result = result.Replace("$GETUPDATEVALUELIST$\r\n", GetUpdateValueList(fields));
            result = result.Replace("$SETORDINALS$\r\n", GetSetOrdinals(fields, isIDRIS, hasIDCode));
            result = result.Replace("$FILLFIELDS$\r\n", GetFillFields(fields));

            return result;
        }

        private static string GetFillFields(List<FieldItem> fields)
        {
            //throw new NotImplementedException();
            StringBuilder result = new StringBuilder();
            foreach (FieldItem currFieldItem in fields)
            {
                result.Append("        obj.");
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
                            result.Append("(int?)");
                            break;
                        case "DECIMAL":
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
                        default:
                            result.Append("###");
                            break;
                    }
                }
                result.Append("dr.");
                switch (currFieldItem.FieldType)
                {
                    case "BIT":
                        result.Append("GetBoolean");
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
                result.Append("        sb.Append(\"");
                if (!firstField)
                {
                    result.Append(", ");
                }
                result.Append("[");
                result.Append(currFieldItem.FieldName);
                result.AppendLine("] = \");");
                result.Append("        sb.Append(");
                switch (currFieldItem.FieldType)
                {
                    case "CHAR":
                    case "NCHAR":
                    case "VARCHAR":
                    case "NVARCHAR":
                        result.Append("SQL.StringToSQLQuoted(obj.");
                        result.Append(currFieldItem.FieldName);
                        result.Append(")");
                        break;
                    default:
                        // ### may need to be split out more to other types
                        result.Append("obj.");
                        result.Append(currFieldItem.FieldName);
                        result.Append("?.ToString() ?? \"NULL\"");
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
                    result.AppendLine("        sb.Append(\", \");");
                }
                result.Append("        sb.Append(");
                switch (currFieldItem.FieldType)
                {
                    case "CHAR":
                    case "NCHAR":
                    case "VARCHAR":
                    case "NVARCHAR":
                        result.Append("SQL.StringToSQLQuoted(obj.");
                        result.Append(currFieldItem.FieldName);
                        result.Append(")");
                        break;
                    default:
                        // ### may need to be split out more to other types
                        result.Append("obj.");
                        result.Append(currFieldItem.FieldName);
                        result.Append("?.ToString() ?? \"NULL\"");
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
                result.Append("        sb.Append(\"");
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
                result.Append("        sb.Append(\", [");
                result.Append(currFieldItem.FieldName);
                result.AppendLine("]\");");
            }
            return result.ToString();
        }

        private static string GetSetOrdinals(List<FieldItem> fields, bool isIDRIS, bool hasIDCode)
        {
            StringBuilder result = new StringBuilder();
            int nextOrdinal;
            if (isIDRIS)
            {
                nextOrdinal = 4; // IDRIS only has 4 header fields
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
                result.Append("        _ord");
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
                result.Append("    public ");
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
                        result.Append("int");
                        if (!currFieldItem.NotNull)
                        {
                            result.Append("?");
                        }
                        break;
                    case "DECIMAL":
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
                result.Append("    private static int _ord");
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