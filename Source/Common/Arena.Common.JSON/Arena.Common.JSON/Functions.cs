// Functions.cs - 10/17/2017

using Arena.Common.Errors;
using System;
using System.Text;

namespace Arena.Common.JSON
{
    static internal class Functions
    {
        internal static int IndentSize = 2;

        internal static void SkipWhitespace(string input, ref int pos)
        {
            while (pos < input.Length && char.IsWhiteSpace(input[pos]))
            {
                pos++;
            }
        }

        /// <summary>
        /// Expects "pos" to be first char of string after initial quote
        /// </summary>
        internal static string GetStringValue(string input, ref int pos)
        {
            StringBuilder value = new StringBuilder();
            char c;
            bool lastWasSlash = false;
            while (pos < input.Length)
            {
                // get next char
                c = input[pos];
                pos++;
                // handle string value
                if (!lastWasSlash && c == '\\') // slashed char
                {
                    lastWasSlash = true;
                    continue;
                }
                if (lastWasSlash) // here is the slashed char
                {
                    lastWasSlash = false;
                    string escapedChar;
                    if (c == 'u')
                    {
                        if (pos + 4 > input.Length)
                        {
                            // doesn't have four hex chars after "u"
                            throw new SystemException(ErrorHandler.FixMessage("Invalid escaped char sequence"));
                        }
                        escapedChar = FromUnicodeChar("\\u" + input[pos] + input[pos + 1] + input[pos + 2] + input[pos + 3]);
                        pos = pos + 4;
                    }
                    else
                    {
                        escapedChar = FromEscapedChar(c);
                    }
                    value.Append(escapedChar);
                    continue;
                }
                if (c == '\"') // end of string
                {
                    return value.ToString();
                }
                // any other char in a string
                value.Append(c);
                continue;
            }
            // incorrect syntax!
            throw new SystemException(ErrorHandler.FixMessage("Incorrect syntax"));
        }

        internal static string ToJsonString(string input)
        {
            // handle escaping of special chars here
            StringBuilder result = new StringBuilder();
            int pos = 0;
            char c;
            while (pos < input.Length)
            {
                c = input[pos];
                pos++;
                if (c == '\\')
                {
                    result.Append("\\\\");
                }
                else if (c == '\"')
                {
                    result.Append("\\\"");
                }
                else if (c == '\r')
                {
                    result.Append("\\r");
                }
                else if (c == '\n')
                {
                    result.Append("\\n");
                }
                else if (c == '\t')
                {
                    result.Append("\\t");
                }
                else if (c == '\b')
                {
                    result.Append("\\b");
                }
                else if (c == '\f')
                {
                    result.Append("\\f");
                }
                else if (c < 32 || c == 127 || c == 129 || c == 141 || c == 143 ||
                         c == 144 || c == 157 || c == 160 || c == 173 || c > 255)
                {
                    // ascii control chars, unused chars, or unicode chars
                    result.Append(string.Format("\\u{0:x4}", (int)c));
                }
                else
                {
                    result.Append(c);
                }
            }
            return result.ToString();
        }

        internal static bool IsNumericType(object value)
        {
            if (value == null)
            {
                return false;
            }
            Type t = value.GetType();
            switch (Type.GetTypeCode(t))
            {
                case TypeCode.Byte:
                case TypeCode.SByte:
                case TypeCode.Int16:
                case TypeCode.Int32:
                case TypeCode.Int64:
                case TypeCode.UInt16:
                case TypeCode.UInt32:
                case TypeCode.UInt64:
                case TypeCode.Single:
                case TypeCode.Double:
                case TypeCode.Decimal:
                    return true;
                case TypeCode.Object:
                    if (t.IsGenericType && t.GetGenericTypeDefinition() == typeof(Nullable<>))
                    {
                        return IsNumericType(Nullable.GetUnderlyingType(t));
                    }
                    return false;
            }
            return false;
        }

        internal static string FromEscapedChar(char c)
        {
            string result;
            switch (c)
            {
                case '"':
                    result = "\"";
                    break;
                case '\\':
                    result = "\\";
                    break;
                case '/':
                    result = "/";
                    break;
                case 'b':
                    result = "\b";
                    break;
                case 'f':
                    result = "\f";
                    break;
                case 'n':
                    result = "\n";
                    break;
                case 'r':
                    result = "\r";
                    break;
                case 't':
                    result = "\t";
                    break;
                default:
                    // escaped unicode (\uXXXX) is handled in FromUnicodeChar()
                    throw new System.Exception(ErrorHandler.FixMessage($"Unknown escaped char: \"\\{c}\""));
            }
            return result;
        }

        internal static string FromUnicodeChar(string value)
        {
            string result;
            // value should be in the exact format "\u####", where # is a hex digit
            if (string.IsNullOrEmpty(value) || value.Length != 6)
            {
                throw new System.Exception(ErrorHandler.FixMessage($"Unknown unicode char: \"{value}\""));
            }
            result = Convert.ToChar(Convert.ToUInt16(value.Substring(2, 4), 16)).ToString();
            return result;
        }

    }
}
