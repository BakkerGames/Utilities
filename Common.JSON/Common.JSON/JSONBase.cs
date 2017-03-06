// JSONBase.cs - 03/04/2017

using System;

namespace Common.JSON
{
    class JSONBase
    {
        public static object FromString(string input)
        {
            object result = null;
            int pos = 0;
            Functions.SkipWhitespace(input, ref pos);
            if (pos < input.Length)
            {
                if (input[pos] == '{')
                {
                    result = JSONObject.FromString(input, ref pos);
                }
                else if (input[pos] == '[')
                {
                    result = JSONArray.FromString(input, ref pos);
                }
                if (result != null)
                {
                    Functions.SkipWhitespace(input, ref pos);
                    if (pos == input.Length)
                    {
                        return result;
                    }
                }
            }
            throw new SystemException("Incorrectly formatted JSONObject or JSONArray");
        }
    }
}
