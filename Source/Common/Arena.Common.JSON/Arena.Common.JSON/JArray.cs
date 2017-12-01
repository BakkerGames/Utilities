// JArray.cs - 11/30/2017

using Arena.Common.Errors;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace Arena.Common.JSON
{
    sealed public class JArray : IEnumerable<object>
    {
        private List<object> _data = new List<object>();

        public IEnumerator<object> GetEnumerator()
        {
            return ((IEnumerable<object>)_data).GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return ((IEnumerable<object>)_data).GetEnumerator();
        }

        public JArray()
        {
        }

        public JArray(JArray values)
        {
            Append(values);
        }

        public void Clear()
        {
            _data.Clear();
        }

        public void Add(object value)
        {
            _data.Add(value);
        }

        public void Append(JArray jarray)
        {
            if (jarray == null)
            {
                return;
            }
            foreach (object value in jarray)
            {
                _data.Add(value);
            }
        }

        public void Remove(int index)
        {
            _data.RemoveAt(index);
        }

        public object GetValue(int index)
        {
            return _data[index];
        }

        public void SetValue(int index, object value)
        {
            _data[index] = value;
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
            sb.Append("[");
            level++;
            bool addComma = false;
            foreach (object obj in this)
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
                else if (obj.GetType() == typeof(DateTime))
                {
                    // datetime converted to ISO 8601 round-trip format "O"
                    sb.Append("\"");
                    sb.Append(((DateTime)obj).ToString("O"));
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
            sb.Append("]");
            return sb.ToString();
        }

        public static bool TryParse(string input, ref JArray result)
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

        public static JArray Parse(string input)
        {
            JArray result = new JArray();
            if (!string.IsNullOrEmpty(input))
            {
                int pos = 0;
                _Parse(result, input, ref pos);
            }
            return result;
        }

        internal static void _Parse(JArray result, string input, ref int pos)
        {
            char c;
            Functions.SkipWhitespace(input, ref pos);
            if (pos >= input.Length || input[pos] != '[') // not a JArray
            {
                throw new SystemException(ErrorHandler.FixMessage($"Not a JArray, char = '{input[pos]}'"));
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
                    throw new SystemException(ErrorHandler.FixMessage("Quote char when not ReadyForValue"));
                }
                // handle other parts of the syntax
                if (c == ',') // after value, before next
                {
                    if (!inValue && !readyForComma)
                    {
                        throw new SystemException(ErrorHandler.FixMessage("Comma char when not InValue or ReadyForComma"));
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
                if (c == ']') // end of JArray
                {
                    if (!readyForValue && !inValue && !readyForComma)
                    {
                        throw new SystemException(ErrorHandler.FixMessage("EndBracket char when not ReadyForValue, InValue, or ReadyForComma"));
                    }
                    if (value.Length > 0) // ignore empty value
                    {
                        _SaveValue(ref result, value.ToString(), inStringValue);
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
                    JObject._Parse(jo, input, ref pos);
                    result.Add(jo);
                    Functions.SkipWhitespace(input, ref pos);
                    readyForComma = true;
                    readyForValue = false;
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
                    _Parse(ja, input, ref pos);
                    result.Add(ja);
                    Functions.SkipWhitespace(input, ref pos);
                    readyForComma = true;
                    readyForValue = false;
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

        private static void _SaveValue(ref JArray obj, string value, bool inStringValue)
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
            else if (int.TryParse(value, out int intValue))
            {
                obj.Add(intValue); // default to int for anything smaller
            }
            else if (long.TryParse(value, out long longValue))
            {
                obj.Add(longValue);
            }
            else if (decimal.TryParse(value, out decimal decimalValue))
            {
                obj.Add(decimalValue);
            }
            else if (double.TryParse(value, out double doubleValue))
            {
                obj.Add(doubleValue);
            }
            else // unknown or non-numeric value
            {
                throw new SystemException(ErrorHandler.FixMessage($"Invalid value = '{value}'"));
            }
        }

        public JArray Clone()
        {
            // returns a new JArray with no references to any existing objects in memory
            return Parse(ToString());
        }
    }
}
