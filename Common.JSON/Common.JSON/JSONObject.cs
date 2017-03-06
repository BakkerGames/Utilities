// JSONObject.cs - 03/05/2017

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace Common.JSON
{
    sealed public class JSONObject : Dictionary<string, object>
    {
        public JSONObject()
        {
        }

        public JSONObject(string input)
        {
            int pos = 0;
            _FromString(this, input, ref pos);
        }

        public string getString(string key)
        {
            if (string.IsNullOrEmpty(key))
            {
                throw new ArgumentNullException();
            }
            if (!ContainsKey(key))
            {
                return null;
            }
            return (string)this[key];
        }

        public bool getBool(string key)
        {
            if (string.IsNullOrEmpty(key))
            {
                throw new ArgumentNullException();
            }
            if (!ContainsKey(key))
            {
                return false;
            }
            return (bool)this[key];
        }

        public int getInt(string key)
        {
            if (string.IsNullOrEmpty(key))
            {
                throw new ArgumentNullException();
            }
            if (!ContainsKey(key))
            {
                return 0;
            }
            return (int)this[key];
        }

        public long getLong(string key)
        {
            if (string.IsNullOrEmpty(key))
            {
                throw new ArgumentNullException();
            }
            if (!ContainsKey(key))
            {
                return 0;
            }
            return (long)this[key];
        }

        public decimal getDecimal(string key)
        {
            if (string.IsNullOrEmpty(key))
            {
                throw new ArgumentNullException();
            }
            if (!ContainsKey(key))
            {
                return 0;
            }
            return (decimal)this[key];
        }

        public JSONObject getJSONObject(string key)
        {
            if (string.IsNullOrEmpty(key))
            {
                throw new ArgumentNullException();
            }
            if (!ContainsKey(key))
            {
                return null;
            }
            return (JSONObject)this[key];
        }

        public JSONArray getJSONArray(string key)
        {
            if (string.IsNullOrEmpty(key))
            {
                throw new ArgumentNullException();
            }
            if (!ContainsKey(key))
            {
                return null;
            }
            return (JSONArray)this[key];
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
            sb.Append("{");
            level++;
            bool addComma = false;
            object obj;
            foreach (KeyValuePair<string, object> keyvalue in this)
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
                sb.Append("\"");
                sb.Append(Functions.ToJSONString(keyvalue.Key));
                sb.Append("\":");
                if (addWhitespace)
                {
                    sb.Append(" ");
                }
                obj = keyvalue.Value; // easier and matches JSONArray code
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
            sb.Append("}");
            return sb.ToString();
        }

        public static JSONObject FromString(string input)
        {
            if (string.IsNullOrEmpty(input))
            {
                return null;
            }
            int pos = 0;
            JSONObject result = new JSONObject();
            _FromString(result, input, ref pos);
            return result;
        }

        internal static void _FromString(JSONObject result, string input, ref int pos)
        {
            char c;
            Functions.SkipWhitespace(input, ref pos);
            if (pos >= input.Length || input[pos] != '{') // not a JSONObject
            {
                throw new SystemException();
            }
            pos++;
            Functions.SkipWhitespace(input, ref pos);
            bool readyForKey = true;
            bool readyForColon = false;
            bool readyForValue = false;
            bool inValue = false;
            bool inStringValue = false;
            bool readyForComma = false;
            StringBuilder key = new StringBuilder();
            StringBuilder value = new StringBuilder();
            while (pos < input.Length)
            {
                // get next char
                c = input[pos];
                pos++;
                // handle key or string value
                if (c == '\"') // beginning of key or string value
                {
                    if (readyForKey)
                    {
                        readyForKey = false;
                        key.Append(Functions.GetStringValue(input, ref pos));
                        Functions.SkipWhitespace(input, ref pos);
                        readyForColon = true;
                        continue;
                    }
                    if (readyForValue)
                    {
                        inValue = true;
                        inStringValue = true;
                        readyForValue = false;
                        value.Append(Functions.GetStringValue(input, ref pos));
                        _SaveKeyValue(ref result, key.ToString(), value.ToString(), inStringValue);
                        Functions.SkipWhitespace(input, ref pos);
                        inValue = false;
                        inStringValue = false;
                        readyForComma = true;
                        key.Clear();
                        value.Clear();
                        continue;
                    }
                    throw new SystemException();
                }
                // handle other parts of the syntax
                if (c == ':') // between key and value
                {
                    if (!readyForColon)
                    {
                        throw new SystemException();
                    }
                    Functions.SkipWhitespace(input, ref pos);
                    readyForValue = true;
                    readyForColon = false;
                    continue;
                }
                if (c == ',') // after value, before next key
                {
                    if (!inValue && !readyForComma)
                    {
                        throw new SystemException();
                    }
                    if (inValue)
                    {
                        _SaveKeyValue(ref result, key.ToString(), value.ToString(), inStringValue);
                    }
                    Functions.SkipWhitespace(input, ref pos);
                    inValue = false;
                    inStringValue = false;
                    readyForComma = false;
                    readyForKey = true;
                    key.Clear();
                    value.Clear();
                    continue;
                }
                if (c == '}') // end of JSONObject
                {
                    if (!readyForKey && !inValue && !readyForComma)
                    {
                        throw new SystemException();
                    }
                    if (key.Length > 0) // ignore empty key
                    {
                        _SaveKeyValue(ref result, key.ToString(), value.ToString(), inStringValue);
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
                    _FromString(jo, input, ref pos);
                    result.Add(key.ToString(), jo);
                    Functions.SkipWhitespace(input, ref pos);
                    readyForComma = true;
                    readyForValue = false;
                    key.Clear();
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
                    JSONArray._FromString(ja, input, ref pos);
                    result.Add(key.ToString(), ja);
                    Functions.SkipWhitespace(input, ref pos);
                    readyForComma = true;
                    readyForValue = false;
                    key.Clear();
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

        private static void _SaveKeyValue(ref JSONObject obj, string key, string value, bool inStringValue)
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
                    obj.Add(key, datetimeValue);
                }
                else
                {
                    obj.Add(key, value);
                }
            }
            else if (value == "null")
            {
                obj.Add(key, null);
            }
            else if (value == "true")
            {
                obj.Add(key, true);
            }
            else if (value == "false")
            {
                obj.Add(key, false);
            }
            else if (int.TryParse(value, out intValue))
            {
                obj.Add(key, intValue); // default to int for anything smaller
            }
            else if (long.TryParse(value, out longValue))
            {
                obj.Add(key, longValue);
            }
            else if (decimal.TryParse(value, out decimalValue))
            {
                obj.Add(key, decimalValue);
            }
            else if (double.TryParse(value, out doubleValue))
            {
                obj.Add(key, doubleValue);
            }
            else // unknown or non-numeric value
            {
                throw new SystemException();
            }
        }
    }
}
