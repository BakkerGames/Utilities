// Program.cs - 11/10/2017

using System;
using System.IO;
using System.Text;

namespace FixSPFuncViewScripts
{
    class Program
    {
        static string _dirName = "";
        static bool _addGo = false;
        static bool _addCRLF = false;

        static int Main(string[] args)
        {
            if (args == null || args.Length == 0)
            {
                Console.WriteLine("Syntax: FixSPFuncViewScripts <path> {/go} {/crlf}");
#if DEBUG
                Console.ReadKey();
#endif
                return (1);
            }
            foreach (string currArg in args)
            {
                if (currArg.StartsWith("/"))
                {
                    if (currArg.Equals("/go", StringComparison.OrdinalIgnoreCase))
                    {
                        _addGo = true;
                    }
                    if (currArg.Equals("/crlf", StringComparison.OrdinalIgnoreCase))
                    {
                        _addCRLF = true;
                    }
                }
                else
                {
                    if (!string.IsNullOrEmpty(_dirName))
                    {
                        Console.WriteLine("Directory already specified");
#if DEBUG
                        Console.ReadKey();
#endif
                        return (2);
                    }
                    if (!Directory.Exists(currArg))
                    {
                        Console.WriteLine($"Directory not found: {currArg}");
#if DEBUG
                        Console.ReadKey();
#endif
                        return (3);
                    }
                    _dirName = currArg;
                }
            }
            DoAllScriptsInPath(_dirName);
            Console.WriteLine("*** Done ***");
#if DEBUG
            Console.ReadKey();
#endif
            return (0);
        }

        private static void DoAllScriptsInPath(string path)
        {
            foreach (string filename in Directory.GetFiles(path, "*.sql"))
            {
                if (!filename.ToLower().EndsWith(".sql"))
                {
                    continue; // could be .sqlproj
                }
                if (filename.ToLower().Contains(".table.sql")
                    || filename.ToLower().Contains("\\tables\\"))
                {
                    continue; // skip table scripts
                }
                DoOneScript(filename);
            }
            foreach (string subPath in Directory.GetDirectories(path))
            {
                if (subPath.Contains("\\."))
                {
                    continue;
                }
                DoAllScriptsInPath(subPath);
            }
        }

        private static void DoOneScript(string filename)
        {
            StringBuilder sb = new StringBuilder();
            StringBuilder origFile = new StringBuilder();
            string lineUC;
            bool skipNextGo;
            bool inMultiLineComment;
            sb.Clear();
            skipNextGo = false;
            inMultiLineComment = false;
            string outLine;
            bool lastWasGo = false;
            bool lastWasBlank = false;
            // look for crlf at end of file
            string tempCRLF = File.ReadAllText(filename);
            bool crlfAtEnd = (tempCRLF.EndsWith("\r") || tempCRLF.EndsWith("\n"));
            // process each line
            foreach (string line in File.ReadAllLines(filename))
            {
                if (origFile.Length > 0)
                {
                    origFile.AppendLine();
                }
                origFile.Append(line);
                outLine = line.TrimEnd();
                lineUC = line.TrimEnd().ToUpper();
                string outLine2 = outLine;
                // remove multiline scripting comments
                if (!inMultiLineComment
                    && lineUC.TrimStart().StartsWith("/*")
                    && lineUC.Contains("==SCRIPTING PARAMETERS=="))
                {
                    inMultiLineComment = true;
                }
                if (inMultiLineComment)
                {
                    if (lineUC.EndsWith("*/"))
                    {
                        inMultiLineComment = false;
                        continue;
                    }
                    else
                    {
                        continue;
                    }
                }
                if (skipNextGo && lineUC.Equals("GO"))
                {
                    skipNextGo = false;
                    continue;
                }
                // remove junk SET statements
                if (lineUC.StartsWith("SET QUOTED_IDENTIFIER") ||
                    lineUC.StartsWith("SET ANSI_NULLS") ||
                    lineUC.StartsWith("SET ANSI_PADDING"))
                {
                    skipNextGo = true;
                    continue;
                }
                // tab expansion and replacement
                if (outLine.Contains("\t"))
                {
                    StringBuilder fixTab = new StringBuilder();
                    foreach (char c in outLine)
                    {
                        if (c == '\t')
                        {
                            fixTab.Append(new string(' ', 4 - (fixTab.Length % 4)));
                        }
                        else
                        {
                            fixTab.Append(c);
                        }
                    }
                    outLine = fixTab.ToString();
                }
                if (outLine.StartsWith("    "))
                {
                    int firstChar = 0;
                    for (int i = 0; i < outLine.Length; i++)
                    {
                        if (outLine[i] != ' ')
                        {
                            firstChar = i;
                            break;
                        }
                    }
                    if (firstChar > 0 && firstChar == (firstChar / 4) * 4)
                    {
                        outLine = $"{new string('\t', firstChar / 4)}{outLine.Substring(firstChar)}";
                    }
                }
                // done with this line
                if (sb.Length > 0)
                {
                    sb.AppendLine();
                }
                sb.Append(outLine);
                lastWasGo = (outLine.Equals("GO", StringComparison.OrdinalIgnoreCase));
                lastWasBlank = string.IsNullOrEmpty(outLine);
            }
            if (_addGo && !lastWasGo)
            {
                if (sb.Length > 0)
                {
                    sb.AppendLine();
                }
                sb.Append("GO");
            }
            if (!lastWasBlank && (crlfAtEnd || _addCRLF))
            {
                sb.AppendLine();
            }
            if (crlfAtEnd)
            {
                origFile.AppendLine();
            }
            // compare new to orig to find changes
            if (!sb.ToString().Equals(origFile.ToString()))
            {
                Console.WriteLine($"{filename} - Changed");
                File.WriteAllText(filename, sb.ToString(), new UTF8Encoding(false, true));
            }
        }
    }
}
