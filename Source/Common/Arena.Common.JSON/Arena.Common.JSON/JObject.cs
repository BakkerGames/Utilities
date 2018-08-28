// JObject.cs - 08/28/2018

using Arena.Common.Errors;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace Arena.Common.JSON
{
    sealed public partial class JObject : IEnumerable<KeyValuePair<string, object>>
    {
        private const string _dateOnlyFormat = "yyyy-MM-dd";
        private const string _dateTimeFormat = "O";

        private Dictionary<string, object> _data = new Dictionary<string, object>();

        public IEnumerator<KeyValuePair<string, object>> GetEnumerator()
        {
            return ((IEnumerable<KeyValuePair<string, object>>)_data).GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return ((IEnumerable<KeyValuePair<string, object>>)_data).GetEnumerator();
        }

        public JObject()
        {
        }

        public JObject(JObject values)
        {
            Append(values);
        }

        public int Count
        {
            get
            {
                return _data.Count;
            }
        }

        public void Clear()
        {
            _data.Clear();
        }

        public void Add(string name, object value)
        {
            _data.Add(name, value);
        }

        public void Append(JObject values)
        {
            if (values == null)
            {
                return;
            }
            foreach (KeyValuePair<string, object> keyvalue in values)
            {
                // will overwrite any matching values
                SetValue(keyvalue.Key, keyvalue.Value);
            }
        }

        public void Remove(string name)
        {
            if (_data.ContainsKey(name))
            {
                _data.Remove(name);
            }
        }

        public object GetValue(string name)
        {
            if (_data.ContainsKey(name))
            {
                return _data[name];
            }
            throw new SystemException(ErrorHandler.FixMessage($"Key not found: {name}"));
        }

        public object GetValueOrNull(string name)
        {
            if (_data.ContainsKey(name))
            {
                return _data[name];
            }
            return null;
        }

        public bool IsNull(string name)
        {
            if (_data.ContainsKey(name))
            {
                return (_data[name] == null);
            }
            return true;
        }

        public void SetValue(string name, object value)
        {
            if (_data.ContainsKey(name))
            {
                _data[name] = value;
            }
            else
            {
                _data.Add(name, value);
            }
        }

        public bool Contains(string name)
        {
            return _data.ContainsKey(name);
        }

        public List<string> Names()
        {
            List<string> result = new List<string>();
            foreach (string key in _data.Keys)
            {
                result.Add(key);
            }
            return result;
        }

        public override string ToString()
        {
            return _ToString(JsonFormat.None, 0);
        }

        public string ToString(JsonFormat format)
        {
            return _ToString(format, 0);
        }

        internal string _ToString(JsonFormat format, int level)
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
                    if (format == JsonFormat.Indent)
                    {
                        sb.AppendLine();
                        sb.Append(new string(' ', level * Functions.IndentSize));
                    }
                }
                else
                {
                    if (format == JsonFormat.Indent)
                    {
                        sb.AppendLine();
                        sb.Append(new string(' ', level * Functions.IndentSize));
                    }
                    addComma = true;
                }
                sb.Append("\"");
                sb.Append(Functions.ToJsonString(keyvalue.Key));
                sb.Append("\":");
                if (format == JsonFormat.Indent)
                {
                    sb.Append(" ");
                }
                obj = keyvalue.Value; // easier and matches JArray code
                if (obj == null)
                {
                    sb.Append("null"); // must be lowercase
                }
                else if (obj.GetType() == typeof(bool))
                {
                    sb.Append((bool)obj ? "true" : "false"); // must be lowercase
                }
                else if (Functions.IsDecimalType(obj))
                {
                    // normalize decimal places
                    sb.Append(Functions.NormalizeDecimal(obj.ToString()));
                }
                else if (Functions.IsNumericType(obj))
                {
                    // number with no quotes
                    sb.Append(obj.ToString());
                }
                else if (obj.GetType() == typeof(DateTime))
                {
                    // datetime converted to string format
                    sb.Append("\"");
                    DateTime tempDT = (DateTime)obj;
                    if (tempDT.Hour + tempDT.Minute + tempDT.Second + tempDT.Millisecond == 0)
                    {
                        sb.Append(tempDT.ToString(_dateOnlyFormat));
                    }
                    else
                    {
                        sb.Append(tempDT.ToString(_dateTimeFormat));
                    }
                    sb.Append("\"");
                }
                else if (obj.GetType() == typeof(JObject))
                {
                    sb.Append(((JObject)obj)._ToString(format, level));
                }
                else if (obj.GetType() == typeof(JArray))
                {
                    sb.Append(((JArray)obj)._ToString(format, level));
                }
                else if (obj.GetType().IsArray)
                {
                    JArray tempArray = new JArray();
                    for (int i = ((Array)obj).GetLowerBound(0); i <= ((Array)obj).GetUpperBound(0); i++)
                    {
                        if (((Array)obj).Rank == 1)
                        {
                            tempArray.Add(((Array)obj).GetValue(i));
                        }
                        else
                        {
                            JArray tempArray2 = new JArray();
                            for (int j = ((Array)obj).GetLowerBound(1); j <= ((Array)obj).GetUpperBound(1); j++)
                            {
                                if (((Array)obj).Rank == 2)
                                {
                                    tempArray2.Add(((Array)obj).GetValue(i, j));
                                }
                                else
                                {
                                    JArray tempArray3 = new JArray();
                                    for (int k = ((Array)obj).GetLowerBound(2); k <= ((Array)obj).GetUpperBound(2); k++)
                                    {
                                        tempArray3.Add(((Array)obj).GetValue(i, j, k));
                                    }
                                    tempArray2.Add(tempArray3);
                                }
                            }
                            tempArray.Add(tempArray2);
                        }
                    }
                    sb.Append(tempArray.ToString());
                }
                else if (obj.GetType().IsGenericType && obj is IEnumerable)
                {
                    JArray tempArray = new JArray();
                    foreach (object o in (IEnumerable)obj)
                    {
                        tempArray.Add(o);
                    }
                    sb.Append(tempArray.ToString());
                }
                else // string or other type which needs quotes
                {
                    sb.Append("\"");
                    sb.Append(Functions.ToJsonString(obj.ToString()));
                    sb.Append("\"");
                }
            }
            level--;
            if (addComma && format == JsonFormat.Indent)
            {
                sb.AppendLine();
                sb.Append(new string(' ', level * Functions.IndentSize));
            }
            sb.Append("}");
            return sb.ToString();
        }

        public static bool TryParse(string input, ref JObject result)
        {
            try
            {
                result = Parse(input);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public static JObject Parse(string input)
        {
            JObject result = new JObject();
            if (!string.IsNullOrEmpty(input))
            {
                int pos = 0;
                _Parse(result, input, ref pos);
            }
            return result;
        }

        internal static void _Parse(JObject result, string input, ref int pos)
        {
            char c;
            Functions.SkipWhitespace(input, ref pos);
            if (pos >= input.Length || input[pos] != '{') // not a JObject
            {
                throw new SystemException(ErrorHandler.FixMessage($"Not a JObject, char = '{input[pos]}'"));
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
                // whitespace is always allowed at this point in the loop
                if (char.IsWhiteSpace(c))
                {
                    Functions.SkipWhitespace(input, ref pos);
                    continue;
                }
                if (c == '/') // ignore comments, //... or /*...*/
                {
                    Functions.SkipWhitespace(input, ref pos);
                    continue;
                }
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
                    throw new SystemException(ErrorHandler.FixMessage("Quote char when not ReadyForKey or ReadyForValue"));
                }
                // handle other parts of the syntax
                if (c == ':') // between key and value
                {
                    if (!readyForColon)
                    {
                        throw new SystemException(ErrorHandler.FixMessage("Colon char when not ReadyForColon"));
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
                        throw new SystemException(ErrorHandler.FixMessage("Comma char when not InValue or ReadyForComma"));
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
                if (c == '}') // end of JObject
                {
                    if (!readyForKey && !inValue && !readyForComma)
                    {
                        throw new SystemException(ErrorHandler.FixMessage("EndBrace char when not ReadyForKey, InValue, or ReadyForComma"));
                    }
                    if (key.Length > 0) // ignore empty key
                    {
                        _SaveKeyValue(ref result, key.ToString(), value.ToString(), inStringValue);
                    }
                    break;
                }
                // handle JObjects and JArrays
                if (c == '{') // JObject as a value
                {
                    if (!readyForValue)
                    {
                        throw new SystemException(ErrorHandler.FixMessage("BeginBrace char when not ReadyForValue"));
                    }
                    pos--;
                    JObject jo = new JObject();
                    _Parse(jo, input, ref pos);
                    result.Add(key.ToString(), jo);
                    Functions.SkipWhitespace(input, ref pos);
                    readyForComma = true;
                    readyForValue = false;
                    key.Clear();
                    value.Clear();
                    continue;
                }
                if (c == '[') // JArray as a value
                {
                    if (!readyForValue)
                    {
                        throw new SystemException(ErrorHandler.FixMessage("BeginBracket char when not ReadyForValue"));
                    }
                    pos--;
                    JArray ja = new JArray();
                    JArray._Parse(ja, input, ref pos);
                    result.Add(key.ToString(), ja);
                    Functions.SkipWhitespace(input, ref pos);
                    readyForComma = true;
                    readyForValue = false;
                    key.Clear();
                    value.Clear();
                    continue;
                }
                // not a string, JObject, JArray value
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
                throw new SystemException(ErrorHandler.FixMessage($"Incorrect syntax, char = '{c}'"));
            }
        }

        private static void _SaveKeyValue(ref JObject obj, string key, string value, bool inStringValue)
        {
            if (!inStringValue)
            {
                value = value.TrimEnd(); // helps with parsing
            }
            if (inStringValue)
            {
                // see if the string is a datetime format
                if (DateTime.TryParse(value, CultureInfo.InvariantCulture,
                                      DateTimeStyles.RoundtripKind, out DateTime datetimeValue))
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
            else if (int.TryParse(value, out int intValue))
            {
                obj.Add(key, intValue); // default to int for anything smaller
            }
            else if (long.TryParse(value, out long longValue))
            {
                obj.Add(key, longValue);
            }
            else if (decimal.TryParse(value, out decimal decimalValue))
            {
                obj.Add(key, decimalValue);
            }
            else if (double.TryParse(value, out double doubleValue))
            {
                obj.Add(key, doubleValue);
            }
            else // unknown or non-numeric value
            {
                throw new SystemException(ErrorHandler.FixMessage($"Invalid value = '{value}'"));
            }
        }

        public JObject Clone()
        {
            // returns a new JObject with no references to any existing objects in memory
            return Parse(ToString());
        }
    }
}
