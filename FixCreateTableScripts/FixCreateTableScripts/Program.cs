// Program.cs - 07/25/2017

using System;
using System.IO;
using System.Text;

namespace FixCreateTableScripts
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args == null || args.Length == 0)
            {
                Console.WriteLine("Syntax: FixCreateTableScripts <path>");
                return;
            }
            StringBuilder sb = new StringBuilder();
            StringBuilder def = new StringBuilder();
            string tableName;
            string fieldName;
            string lineUC;
            bool hasChanges;
            bool madeChanges;
            bool inTable;
            bool pastTable;
            bool pastAlter;
            int posStart;
            int posEnd;
            foreach (string filename in Directory.GetFiles(args[0], "*.Table.sql"))
            {
                sb.Clear();
                def.Clear();
                hasChanges = false;
                madeChanges = false;
                inTable = false;
                pastTable = false;
                pastAlter = false;
                string defLine;
                string outLine;
                tableName = "";
                foreach (string line in File.ReadAllLines(filename))
                {
                    outLine = line.TrimEnd();
                    lineUC = line.TrimEnd().ToUpper();
                    // fix standard issues with create table scripts
                    if (string.IsNullOrEmpty(lineUC))
                    {
                        hasChanges = true;
                        continue;
                    }
                    posStart = lineUC.IndexOf("SET QUOTED_IDENTIFIER OFF");
                    if (lineUC == "SET QUOTED_IDENTIFIER OFF")
                    {
                        lineUC = "SET QUOTED_IDENTIFIER ON";
                        outLine = lineUC;
                        hasChanges = true;
                    }
                    if (lineUC.Contains(", FILLFACTOR = 90") || lineUC.Contains(", FILLFACTOR = 95"))
                    {
                        posStart = lineUC.IndexOf(", FILLFACTOR = ");
                        lineUC = lineUC.Substring(0, posStart) + lineUC.Substring(posStart + ", FILLFACTOR = 9x".Length);
                        outLine = outLine.Substring(0, posStart) + outLine.Substring(posStart + ", FILLFACTOR = 9x".Length);
                        hasChanges = true;
                    }
                    if (lineUC.Contains(" TEXTIMAGE_ON [PRIMARY]"))
                    {
                        posStart = lineUC.IndexOf(" TEXTIMAGE_ON [PRIMARY]");
                        lineUC = lineUC.Substring(0, posStart) + lineUC.Substring(posStart + " TEXTIMAGE_ON [PRIMARY]".Length);
                        outLine = outLine.Substring(0, posStart) + outLine.Substring(posStart + " TEXTIMAGE_ON [PRIMARY]".Length);
                        hasChanges = true;
                    }
                    // check for inline defaults instead of alter table defaults
                    if (!inTable && !pastTable && !pastAlter)
                    {
                        posStart = lineUC.IndexOf("CREATE TABLE");
                        posEnd = lineUC.IndexOf("(");
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
                            defLine = line.Substring(posStart, posEnd - posStart);
                            outLine = line.Substring(0, posStart) + line.Substring(posEnd);
                            // check for single parens around default
                            if (defLine.Contains("(") && !defLine.Contains("((") &&
                                defLine.Contains(")") && !defLine.Contains("))"))
                            {
                                defLine = defLine.Replace("(", "((").Replace(")", "))");
                            }
                            // build new alter table section
                            def.Append("ALTER TABLE ");
                            def.Append(tableName);
                            def.Append(" ADD ");
                            def.Append(defLine);
                            def.Append(" FOR ");
                            def.AppendLine(fieldName);
                            def.AppendLine("GO");
                            hasChanges = true;
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
                            hasChanges = true;
                        }
                    }
                    sb.AppendLine(outLine);
                }
                if (hasChanges)
                {
                    if (!madeChanges)
                    {
                        sb.Append(def.ToString());
                        madeChanges = true;
                    }
                    Console.WriteLine($"{filename} - Changed");
                    File.WriteAllText(filename, sb.ToString(), Encoding.UTF8);
                }
            }
            Console.WriteLine("*** Done ***");
#if DEBUG
            Console.ReadKey();
#endif
        }
    }
}
