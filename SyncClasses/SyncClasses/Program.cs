// Program.cs - 11/10/2017

// This is a little utility to copy generated data classes from a single directory
// to all projects in a source code folder. It will replace the token "$NAMESPACE$"
// with the target parent folder name individually. Files are skipped if there are
// no changes.

// Syntax: SyncClasses <Class File Folder> <Source Code Root Folder>

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace SyncClasses
{
    class Program
    {
        private static List<string> fromFiles = new List<string>();

        static void Main(string[] args)
        {
            string fromPath = args[0];
            string toPath = args[1];

            Console.WriteLine($"Copying from \"{fromPath}\" to \"{toPath}\"");
            Console.WriteLine();

            foreach (string fromFilename in Directory.GetFiles(fromPath, "*.cs"))
            {
                string baseFilename = fromFilename.Substring(fromFilename.LastIndexOf("\\") + 1);
                fromFiles.Add(baseFilename);
            }

            FindMatching(fromPath, toPath);

#if DEBUG
            Console.WriteLine("Press enter to continue...");
            Console.ReadLine();
#endif
        }

        static void FindMatching(string fromPath, string currPath)
        {
            // check for matching files in this directory
            foreach (string currFilename in Directory.GetFiles(currPath, "*.cs"))
            {
                string baseFilename = currFilename.Substring(currFilename.LastIndexOf("\\") + 1);
                if (fromFiles.Contains(baseFilename))
                {
                    //Console.WriteLine(baseFilename);
                    SyncClass(baseFilename, fromPath, currPath);
                }
            }

            // now traverse all subdirectories
            foreach (string subDirName in Directory.GetDirectories(currPath))
            {
                FindMatching(fromPath, subDirName);
            }
        }

        static void SyncClass(string baseFilename, string fromPath, string toPath)
        {
            string fromFileText = File.ReadAllText($"{fromPath}\\{baseFilename}");
            string toFileText = File.ReadAllText($"{toPath}\\{baseFilename}");
            // fix namespace to match project name, assumes directory name = project name
            string projectName = toPath.Substring(toPath.LastIndexOf("\\") + 1);
            fromFileText = fromFileText.Replace("$NAMESPACE$", projectName);
            // only copy files if different
            if (!fromFileText.Equals(toFileText))
            {
                Console.WriteLine($"{toPath}\\{baseFilename}");
                File.WriteAllText($"{toPath}\\{baseFilename}", fromFileText, new UTF8Encoding(false, true));
            }
        }
    }
}
