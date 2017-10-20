// Program.cs - 10/20/2017

using System;
using System.IO;
using System.Text;

namespace FixSPFuncViewScripts
{
    class Program
    {
        static int Main(string[] args)
        {
            if (args == null || args.Length == 0 || !Directory.Exists(args[0]))
            {
                Console.WriteLine("Syntax: FixSPFuncViewScripts <path>");
#if DEBUG
                Console.ReadKey();
#endif
                return (1);
            }
            DoAllScriptsInPath(args[0]);
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
            string lineUC;
            bool hasChanges;
            bool skipNextGo;
            bool inMultiLineComment;
            sb.Clear();
            hasChanges = false;
            skipNextGo = false;
            inMultiLineComment = false;
            string outLine;
            foreach (string line in File.ReadAllLines(filename))
            {
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
                        hasChanges = true;
                        continue;
                    }
                    else
                    {
                        hasChanges = true;
                        continue;
                    }
                }
                // fix standard issues with create table scripts
                if (string.IsNullOrEmpty(lineUC))
                {
                    hasChanges = true;
                    continue;
                }
                if (skipNextGo && lineUC.Equals("GO"))
                {
                    skipNextGo = false;
                    hasChanges = true;
                    continue;
                }
                // remove junk SET statements
                if (lineUC.StartsWith("SET QUOTED_IDENTIFIER") ||
                    lineUC.StartsWith("SET ANSI_NULLS") ||
                    lineUC.StartsWith("SET ANSI_PADDING"))
                {
                    skipNextGo = true;
                    hasChanges = true;
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
                if (!outLine2.Equals(outLine))
                {
                    hasChanges = true;
                }
                // done with this line
                sb.AppendLine(outLine);
            }
            if (hasChanges)
            {
                Console.WriteLine($"{filename} - Changed");
                File.WriteAllText(filename, sb.ToString(), Encoding.UTF8);
            }
        }
    }
}
