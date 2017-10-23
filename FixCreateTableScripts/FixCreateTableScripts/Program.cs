// Program.cs - 10/23/2017

using System;
using System.IO;
using System.Text;

namespace FixCreateTableScripts
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
                Console.WriteLine("Syntax: FixCreateTableScripts <path> {/go} {/crlf}");
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
                    DoOneScript(filename);
                }
            }
            foreach (string subPath in Directory.GetDirectories(path))
            {
                DoAllScriptsInPath(subPath);
            }
        }

        private static void DoOneScript(string filename)
        {
            StringBuilder sb = new StringBuilder();
            StringBuilder def = new StringBuilder();
            StringBuilder origFile = new StringBuilder();
            string tableName;
            string fieldName;
            string lineUC;
            bool madeChanges;
            bool inTable;
            bool pastTable;
            bool pastAlter;
            bool skipNextGo;
            bool inMultiLineComment;
            int posStart;
            int posEnd;
            sb.Clear();
            def.Clear();
            madeChanges = false;
            inTable = false;
            pastTable = false;
            pastAlter = false;
            skipNextGo = false;
            inMultiLineComment = false;
            string defLine;
            string outLine;
            tableName = "";
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
                // fix standard issues with create table scripts
                if (string.IsNullOrEmpty(lineUC))
                {
                    continue;
                }
                if (skipNextGo && lineUC.Equals("GO"))
                {
                    skipNextGo = false;
                    continue;
                }
                if (lineUC.Contains(", FILLFACTOR = 90") || lineUC.Contains(", FILLFACTOR = 95"))
                {
                    posStart = lineUC.IndexOf(", FILLFACTOR = ");
                    lineUC = lineUC.Substring(0, posStart) + lineUC.Substring(posStart + ", FILLFACTOR = 9x".Length);
                    outLine = outLine.Substring(0, posStart) + outLine.Substring(posStart + ", FILLFACTOR = 9x".Length);
                }
                if (lineUC.Contains(" WITH (FILLFACTOR = 90)") || lineUC.Contains(" WITH (FILLFACTOR = 95)"))
                {
                    posStart = lineUC.IndexOf(" WITH (FILLFACTOR = ");
                    lineUC = lineUC.Substring(0, posStart) + lineUC.Substring(posStart + " WITH (FILLFACTOR = 9x)".Length);
                    outLine = outLine.Substring(0, posStart) + outLine.Substring(posStart + " WITH (FILLFACTOR = 9x)".Length);
                }
                if (lineUC.Contains(" TEXTIMAGE_ON [PRIMARY]"))
                {
                    posStart = lineUC.IndexOf(" TEXTIMAGE_ON [PRIMARY]");
                    lineUC = lineUC.Substring(0, posStart) + lineUC.Substring(posStart + " TEXTIMAGE_ON [PRIMARY]".Length);
                    outLine = outLine.Substring(0, posStart) + outLine.Substring(posStart + " TEXTIMAGE_ON [PRIMARY]".Length);
                }
                if (lineUC.Contains(" ENABLE TRIGGER "))
                {
                    skipNextGo = true;
                    continue;
                }
                if (lineUC.Contains(" DISABLE TRIGGER "))
                {
                    skipNextGo = true;
                    continue;
                }
                if (lineUC.Contains(" WITH NOCHECK "))
                {
                    int pos = lineUC.IndexOf(" WITH NOCHECK ");
                    lineUC = lineUC.Substring(0, pos + 6) + lineUC.Substring(pos + 8);
                    outLine = outLine.Substring(0, pos + 6) + outLine.Substring(pos + 8);
                }
                // check for inline defaults instead of alter table defaults
                if (!inTable && !pastTable && !pastAlter)
                {
                    posStart = lineUC.IndexOf("CREATE TABLE");
                    posEnd = lineUC.IndexOf("(");
                    if (posEnd < 0)
                    {
                        posEnd = lineUC.Length;
                    }
                    if (posStart >= 0)
                    {
                        posStart += "CREATE TABLE".Length;
                        tableName = line.Substring(posStart, posEnd - posStart).Trim();
                        inTable = true;
                    }
                }
                else if (inTable && !pastTable && !pastAlter)
                {
                    if (line == "GO")
                    {
                        inTable = false;
                        pastTable = true;
                    }
                    else if (lineUC.Contains(" DEFAULT "))
                    {
                        defLine = null;
                        posStart = line.IndexOf("[");
                        posEnd = line.IndexOf("]");
                        fieldName = line.Substring(posStart, posEnd - posStart + 1);
                        if (lineUC.Contains(" CONSTRAINT "))
                        {
                            posStart = lineUC.IndexOf(" CONSTRAINT ");
                        }
                        else
                        {
                            posStart = lineUC.IndexOf(" DEFAULT ");
                        }
                        posEnd = lineUC.LastIndexOf(",");
                        if (posEnd < 0)
                        {
                            posEnd = line.Length;
                        }
                        if (lineUC.LastIndexOf(" NULL") > posStart && posEnd > lineUC.LastIndexOf(" NULL"))
                        {
                            posEnd = lineUC.LastIndexOf(" NULL");
                        }
                        if (lineUC.LastIndexOf(" NOT NULL") > posStart && posEnd > lineUC.LastIndexOf(" NOT NULL"))
                        {
                            posEnd = lineUC.LastIndexOf(" NOT NULL");
                        }
                        defLine = line.Substring(posStart, posEnd - posStart);
                        outLine = line.Substring(0, posStart) + line.Substring(posEnd);
                        // check for single parens around default
                        if (defLine.Contains("(") && !defLine.Contains("((") &&
                            defLine.Contains(")") && !defLine.Contains("))"))
                        {
                            defLine = defLine.Replace("(", "((").Replace(")", "))");
                        }
                        // build new alter table section
                        if (def.Length > 0)
                        {
                            def.AppendLine();
                        }
                        def.Append("ALTER TABLE ");
                        def.Append(tableName);
                        def.Append(" ADD ");
                        def.Append(defLine);
                        def.Append(" FOR ");
                        def.Append(fieldName);
                        def.AppendLine();
                        def.Append("GO");
                    }
                }
                else if (!pastAlter)
                {
                    if (lineUC.StartsWith("/*") ||
                        lineUC.StartsWith("ALTER") ||
                        lineUC.StartsWith("SET ANSI_NULLS") ||
                        lineUC.StartsWith("SET QUOTED_IDENTIFIER") ||
                        lineUC.StartsWith("CREATE TRIGGER"))
                    {
                        if (sb.Length > 0 && def.Length > 0)
                        {
                            sb.AppendLine();
                        }
                        sb.Append(def.ToString());
                        pastAlter = true;
                        madeChanges = true;
                    }
                }
                if (pastAlter && lineUC.Contains("ALTER TABLE") && lineUC.Contains(" DEFAULT "))
                {
                    if (outLine.Contains("(") && !outLine.Contains("((") &&
                        outLine.Contains(")") && !outLine.Contains("))"))
                    {
                        outLine = outLine.Replace("(", "((").Replace(")", "))");
                    }
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
                outLine = outLine.TrimEnd();
                // done with this line
                if (sb.Length > 0)
                {
                    sb.AppendLine();
                }
                sb.Append(outLine);
                lastWasGo = (outLine.Equals("GO", StringComparison.OrdinalIgnoreCase));
                lastWasBlank = string.IsNullOrEmpty(outLine);
            }
            if (!madeChanges)
            {
                if (sb.Length > 0 && def.Length > 0)
                {
                    sb.AppendLine();
                }
                sb.Append(def.ToString());
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
                File.WriteAllText(filename, sb.ToString(), Encoding.UTF8);
            }
        }
    }
}
