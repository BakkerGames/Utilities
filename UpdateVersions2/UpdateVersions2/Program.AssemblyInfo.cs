// Program.AssemblyInfo.cs - 07/14/2016

using System;
using System.IO;
using System.Text;

namespace UpdateVersions2
{
    partial class Program
    {
        private static string GetAssemblyInfoVersion(string currinfofile)
        {
            string[] infolines;
            infolines = File.ReadAllLines(currinfofile, Encoding.UTF8);
            foreach (string currline in infolines)
            {
                if (currline.Trim().StartsWith("'") |
                    currline.Trim().StartsWith("//"))
                {
                    continue;
                }
                if (currline.IndexOf("Assembly: AssemblyVersion(\"", comp_ic) >= 0)
                {
                    return Functions.SubstringBetween(currline, "Assembly: AssemblyVersion(\"", "\"");
                }
            }
            throw new SystemException("Version not found!");
        }

        private static void UpdateAssemblyInfoVersion(string currinfofile, string newversion)
        {
            string[] infolines;
            infolines = File.ReadAllLines(currinfofile, Encoding.UTF8);
            bool changed = false;
            StringBuilder result = new StringBuilder();
            foreach (string currline in infolines)
            {
                if (result.Length > 0)
                    result.AppendLine();
                if (currline.Trim().StartsWith("'") |
                    currline.Trim().StartsWith("//"))
                {
                    result.Append(currline);
                    continue;
                }
                if (currline.IndexOf("Assembly: AssemblyVersion(\"", comp_ic) >= 0)
                {
                    result.Append(Functions.ReplaceBetween(currline, "Assembly: AssemblyVersion(\"", "\"", newversion));
                    changed = true;
                    continue;
                }
                if (currline.IndexOf("Assembly: AssemblyFileVersion(\"", comp_ic) >= 0)
                {
                    result.Append(Functions.ReplaceBetween(currline, "Assembly: AssemblyFileVersion(\"", "\"", newversion));
                    changed = true;
                    continue;
                }
                // no change needed
                result.Append(currline);
            }
            // ApplicationInfo files end with a CRLF
            result.AppendLine();
            // check if anything changed
            if (changed)
            {
                // write out the changes
                if (File.Exists(currinfofile))
                {
                    File.SetAttributes(currinfofile, FileAttributes.Normal);
                    File.Delete(currinfofile);
                }
                File.WriteAllText(currinfofile, result.ToString());
            }
        }
    }
}
