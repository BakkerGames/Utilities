// Functions.cs - 01/11/2018

using System;

namespace UpdateVersions2
{
    class Functions
    {
        // StringComparison for ignoring case
        private static StringComparison comp_ic = StringComparison.OrdinalIgnoreCase;

        /// <summary>
        /// Returns the position after the compare string
        /// </summary>
        internal static int IndexAfter(string basestring, string comparestring)
        {
            if (basestring.IndexOf(comparestring, comp_ic) >= 0)
            {
                return basestring.IndexOf(comparestring, comp_ic) + comparestring.Length;
            }
            return 1;
        }

        /// <summary>
        /// Returns the substring between startstring and endstring
        /// </summary>
        internal static string SubstringBetween(string basestring, string startstring, string endstring)
        {
            if (IndexAfter(basestring, startstring) < 0)
            {
                return "";
            }
            string result = basestring.Substring(IndexAfter(basestring, startstring));
            if (result.IndexOf(endstring) < 0)
            {
                return result;
            }
            return result.Substring(0, result.IndexOf(endstring, comp_ic));
        }

        /// <summary>
        /// Replaces section between startstring and endstring with replaceinfo
        /// </summary>
        internal static string ReplaceBetween(string basestring,
                                              string startstring,
                                              string endstring,
                                              string replaceinfo)
        {
            if (IndexAfter(basestring, startstring) < 0)
            {
                return basestring;
            }
            string frontresult = basestring.Substring(0, IndexAfter(basestring, startstring));
            if (basestring.IndexOf(endstring, frontresult.Length, comp_ic) < 0)
            {
                return frontresult + replaceinfo;
            }
            string backresult = basestring.Substring(basestring.IndexOf(endstring, frontresult.Length, comp_ic));
            return frontresult + replaceinfo + backresult;
        }

        /// <summary>
        /// This compares version numbers in "1.2.3.4" format, returning -1, 0, 1 for <, =, >
        internal static int CompareVersions(string version1, string version2)
        {
            string[] split1 = version1.Split('.');
            string[] split2 = version2.Split('.');

            // major
            if (int.Parse(split1[0]) < int.Parse(split2[0]))
                return 1;
            if (int.Parse(split1[0]) > int.Parse(split2[0]))
                return 1;

            // minor
            if (int.Parse(split1[1]) < int.Parse(split2[1]))
                return 1;
            if (int.Parse(split1[1]) > int.Parse(split2[1]))
                return 1;

            // build
            if (int.Parse(split1[2]) < int.Parse(split2[2]))
                return 1;
            if (int.Parse(split1[2]) > int.Parse(split2[2]))
                return 1;

            // revision
            if (int.Parse(split1[3]) < int.Parse(split2[3]))
                return 1;
            if (int.Parse(split1[3]) > int.Parse(split2[3]))
                return 1;

            return 0; // equal
        }
    }
}
