// Program.cs - 10/19/2017

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
                if (path.ToLower().Contains(".table.sql")
                    || path.ToLower().Contains("\\tables\\"))
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
            // int posStart;
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
                //if (lineUC.Contains(", FILLFACTOR = 90") || lineUC.Contains(", FILLFACTOR = 95"))
                //{
                //    posStart = lineUC.IndexOf(", FILLFACTOR = ");
                //    lineUC = lineUC.Substring(0, posStart) + lineUC.Substring(posStart + ", FILLFACTOR = 9x".Length);
                //    outLine = outLine.Substring(0, posStart) + outLine.Substring(posStart + ", FILLFACTOR = 9x".Length);
                //    hasChanges = true;
                //}
                //if (lineUC.Contains(" TEXTIMAGE_ON [PRIMARY]"))
                //{
                //    posStart = lineUC.IndexOf(" TEXTIMAGE_ON [PRIMARY]");
                //    lineUC = lineUC.Substring(0, posStart) + lineUC.Substring(posStart + " TEXTIMAGE_ON [PRIMARY]".Length);
                //    outLine = outLine.Substring(0, posStart) + outLine.Substring(posStart + " TEXTIMAGE_ON [PRIMARY]".Length);
                //    hasChanges = true;
                //}
                //if (lineUC.Contains(" WITH NOCHECK "))
                //{
                //    int pos = lineUC.IndexOf(" WITH NOCHECK ");
                //    lineUC = lineUC.Substring(0, pos + 6) + lineUC.Substring(pos + 8);
                //    outLine = outLine.Substring(0, pos + 6) + outLine.Substring(pos + 8);
                //    hasChanges = true;
                //}
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
                //while (outLine.Contains("\t"))
                //{
                //    posStart = outLine.IndexOf('\t');
                //    outLine = $"{outLine.Substring(0, posStart)}{new string(' ', 4 - (posStart % 4))}{outLine.Substring(posStart + 1)}";
                //}
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
