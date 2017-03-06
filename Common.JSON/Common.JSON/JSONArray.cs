// JSONArray.cs - 03/05/2017

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace Common.JSON
{
    sealed public class JSONArray : List<object>
    {
        public JSONArray()
        {
        }

        public JSONArray(string input)
        {
            int pos = 0;
            _FromString(this, input, ref pos);
        }

        public override string ToString()
        {
            return ToString(0, false);
        }

        public string ToString(bool addWhitespace)
        {
            return ToString(0, addWhitespace);
        }

        internal string ToString(int level, bool addWhitespace)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("[");
            level++;
            bool addComma = false;
            foreach (object obj in this)
            {
                if (addComma)
                {
                    sb.Append(",");
                    if (addWhitespace)
                    {
                        sb.AppendLine();
                        sb.Append(new string(' ', level * Functions.IndentSize));
                    }
                }
                else
                {
                    if (addWhitespace)
                    {
                        sb.AppendLine();
                        sb.Append(new string(' ', level * Functions.IndentSize));
                    }
                    addComma = true;
                }
                if (obj == null)
                {
                    sb.Append("null"); // must be lowercase
                }
                else if (obj.GetType() == typeof(bool))
                {
                    sb.Append((bool)obj ? "true" : "false"); // must be lowercase
                }
                else if (Functions.IsNumericType(obj))
                {
                    // number with no quotes
                    sb.Append(obj.ToString());
                }
                else if (obj.GetType() == typeof(JSONObject))
                {
                    sb.Append(((JSONObject)obj).ToString(level, addWhitespace));
                }
                else if (obj.GetType() == typeof(JSONArray))
                {
                    sb.Append(((JSONArray)obj).ToString(level, addWhitespace));
                }
                else if (obj.GetType() == typeof(DateTime))
                {
                    // datetime converted to ISO 8601 round-trip format "O"
                    sb.Append("\"");
                    sb.Append(((DateTime)obj).ToString("O"));
                    sb.Append("\"");
                }
                else // string or other type which needs quotes
                {
                    sb.Append("\"");
                    sb.Append(Functions.ToJSONString(obj.ToString()));
                    sb.Append("\"");
                }
            }
            level--;
            if (addComma && addWhitespace)
            {
                sb.AppendLine();
                sb.Append(new string(' ', level * Functions.IndentSize));
            }
            sb.Append("]");
            return sb.ToString();
        }

        public static JSONArray FromString(string input)
        {
            if (string.IsNullOrEmpty(input))
            {
                return null;
            }
            int pos = 0;
            JSONArray result = new JSONArray();
            _FromString(result, input, ref pos);
            return result;
        }

        internal static void _FromString(JSONArray result, string input, ref int pos)
        {
            char c;
            Functions.SkipWhitespace(input, ref pos);
            if (pos >= input.Length || input[pos] != '[') // not a JSONArray
            {
                throw new SystemException();
            }
            pos++;
            Functions.SkipWhitespace(input, ref pos);
            bool readyForValue = true;
            bool inValue = false;
            bool inStringValue = false;
            bool readyForComma = false;
            StringBuilder value = new StringBuilder();
            while (pos < input.Length)
            {
                // get next char
                c = input[pos];
                pos++;
                // handle string value
                if (c == '\"') // beginning of string value
                {
                    if (readyForValue)
                    {
                        inValue = true;
                        inStringValue = true;
                        readyForValue = false;
                        value.Append(Functions.GetStringValue(input, ref pos));
                        _SaveValue(ref result, value.ToString(), inStringValue);
                        Functions.SkipWhitespace(input, ref pos);
                        inValue = false;
                        inStringValue = false;
                        readyForComma = true;
                        value.Clear();
                        continue;
                    }
                    throw new SystemException();
                }
                // handle other parts of the syntax
                if (c == ',') // after value, before next
                {
                    if (!inValue && !readyForComma)
                    {
                        throw new SystemException();
                    }
                    if (inValue)
                    {
                        _SaveValue(ref result, value.ToString(), inStringValue);
                    }
                    Functions.SkipWhitespace(input, ref pos);
                    inValue = false;
                    inStringValue = false;
                    readyForComma = false;
                    readyForValue = true;
                    value.Clear();
                    continue;
                }
                if (c == ']') // end of JSONArray
                {
                    if (!readyForValue && !inValue && !readyForComma)
                    {
                        throw new SystemException();
                    }
                    if (value.Length > 0) // ignore empty value
                    {
                        _SaveValue(ref result, value.ToString(), inStringValue);
                    }
                    break;
                }
                // handle JSONObjects and JSONArrays
                if (c == '{') // JSONObject as a value
                {
                    if (!readyForValue)
                    {
                        throw new SystemException();
                    }
                    pos--;
                    JSONObject jo = new JSONObject();
                    JSONObject._FromString(jo, input, ref pos);
                    result.Add(jo);
                    Functions.SkipWhitespace(input, ref pos);
                    readyForComma = true;
                    readyForValue = false;
                    value.Clear();
                    continue;
                }
                if (c == '[') // JSONArray as a value
                {
                    if (!readyForValue)
                    {
                        throw new SystemException();
                    }
                    pos--;
                    JSONArray ja = new JSONArray();
                    _FromString(ja, input, ref pos);
                    result.Add(ja);
                    Functions.SkipWhitespace(input, ref pos);
                    readyForComma = true;
                    readyForValue = false;
                    value.Clear();
                    continue;
                }
                // not a string, JSONObject, JSONArray value
                if (readyForValue)
                {
                    readyForValue = false;
                    inValue = true;
                    // don't continue, drop through
                }
                if (inValue)
                {
                    value.Append(c);
                    continue;
                }
                // incorrect syntax!
                throw new SystemException();
            }
        }

        private static void _SaveValue(ref JSONArray obj, string value, bool inStringValue)
        {
            int intValue;
            long longValue;
            decimal decimalValue;
            double doubleValue;
            DateTime datetimeValue;
            if (!inStringValue)
            {
                value = value.TrimEnd(); // helps with parsing
            }
            if (inStringValue)
            {
                // see if the string is a datetime format
                if (DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out datetimeValue))
                {
                    obj.Add(datetimeValue);
                }
                else
                {
                    obj.Add(value);
                }
            }
            else if (value == "null")
            {
                obj.Add(null);
            }
            else if (value == "true")
            {
                obj.Add(true);
            }
            else if (value == "false")
            {
                obj.Add(false);
            }
            else if (int.TryParse(value, out intValue))
            {
                obj.Add(intValue); // default to int for anything smaller
            }
            else if (long.TryParse(value, out longValue))
            {
                obj.Add(longValue);
            }
            else if (decimal.TryParse(value, out decimalValue))
            {
                obj.Add(decimalValue);
            }
            else if (double.TryParse(value, out doubleValue))
            {
                obj.Add(doubleValue);
            }
            else // unknown or non-numeric value
            {
                throw new SystemException();
            }
        }
    }
}
