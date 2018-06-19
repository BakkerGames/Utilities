// Functions.cs - 06/06/2018

// 06/06/2018 - SBakker
//            - Added function NormalizeDecimal() to return deterministic string values
//              for decimal inputs.
// 05/31/2018 - SBakker
//            - Skip non-standard comments as whitespace. Only used for reading,
//              so can comment out lines in config.json files or whatever. Ignores
//              any text in // to eol and /* to */. If end of file is found, that
//              is fine and just ends the comment. This breaks the implementation
//              by not throwing an error if the comments are found, but that seems
//              a reasonable trade-off.

using Arena.Common.Errors;
using System;
using System.Globalization;
using System.Text;

namespace Arena.Common.JSON
{
    static internal class Functions
    {
        internal static int IndentSize = 2;
        internal static CultureInfo decimalCulture = CultureInfo.CreateSpecificCulture("en-US");

        internal static void SkipWhitespace(string input, ref int pos)
        {
            while (true)
            {
                if (pos >= input.Length)
                {
                    break;
                }
                // use c# definition of whitespace
                if (char.IsWhiteSpace(input[pos]))
                {
                    pos++;
                    continue;
                }
                // allow non-standard comment, // to eol
                if (pos + 1 < input.Length && input[pos] == '/' && input[pos + 1] == '/')
                {
                    pos = pos + 2;
                    while (pos < input.Length)
                    {
                        if (input[pos] == '\r' || input[pos] == '\n')
                        {
                            pos++;
                            break;
                        }
                        pos++;
                    }
                    continue;
                }
                // allow non-standard comment, /* to */
                if (pos + 1 < input.Length && input[pos] == '/' && input[pos + 1] == '*')
                {
                    pos = pos + 2;
                    while (pos < input.Length)
                    {
                        if (input[pos - 1] == '*' && input[pos] == '/')
                        {
                            pos++;
                            break;
                        }
                        pos++;
                    }
                    continue;
                }
                break;
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

        internal static bool IsDecimalType(object value)
        {
            if (value == null)
            {
                return false;
            }
            Type t = value.GetType();
            switch (Type.GetTypeCode(t))
            {
                case TypeCode.Single:
                case TypeCode.Double:
                case TypeCode.Decimal:
                    return true;
                case TypeCode.Object:
                    if (t.IsGenericType && t.GetGenericTypeDefinition() == typeof(Nullable<>))
                    {
                        return IsDecimalType(Nullable.GetUnderlyingType(t));
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
                    throw new SystemException(ErrorHandler.FixMessage($"Unknown escaped char: \"\\{c}\""));
            }
            return result;
        }

        internal static string FromUnicodeChar(string value)
        {
            string result;
            // value should be in the exact format "\u####", where # is a hex digit
            if (string.IsNullOrEmpty(value) || value.Length != 6)
            {
                throw new SystemException(ErrorHandler.FixMessage($"Unknown unicode char: \"{value}\""));
            }
            result = Convert.ToChar(Convert.ToUInt16(value.Substring(2, 4), 16)).ToString();
            return result;
        }

        internal static string NormalizeDecimal(string value)
        {
            decimal tempValue;
            if (string.IsNullOrEmpty(value))
            {
                return value;
            }
            if (!decimal.TryParse(value, out tempValue))
            {
                throw new SystemException(ErrorHandler.FixMessage($"Cannot parse as decimal: \"{value}\""));
            }
            return NormalizeDecimal(tempValue);
        }

        internal static string NormalizeDecimal(decimal value)
        {
            string result;
            result = value.ToString(decimalCulture);
            if (result.IndexOf('.') >= 0)
            {
                while (result.Length > 0)
                {
                    if (result.EndsWith(".")) // remove trailing decimal point
                    {
                        result = result.Substring(0, result.Length - 1);
                        break; // done
                    }
                    if (result.EndsWith("0")) // remove trailing zero decimal digits
                    {
                        result = result.Substring(0, result.Length - 1);
                        continue;
                    }
                    break;
                }
            }
            return result;
        }
    }
}
