// Program.Save.cs - 08/16/2016

// 08/16/2016 - SBakker
//            - Removed the saving of version numbers. They have no meaning when using Git.
// 07/29/2016 - SBakker
//            - Skip directory names starting with "."
//            - Don't exclude "Test_" projects during Save.
// 07/20/2016 - SBakker
//            - Added savecount to SaveProgramVersions(), for blank line management.

using System;
using System.IO;
using System.Text;

namespace UpdateVersions2
{
    partial class Program
    {
//        private static int SaveProgramVersions(DirectoryInfo thisdir)
//        {
//            int savecount = 0;
//            string[] filelines;
//            // loop through all the files
//            foreach (FileInfo currfile in thisdir.GetFiles())
//            {
//                if (currfile.FullName.IndexOf(".vbproj", comp_ic) != currfile.FullName.Length - 7 &
//                    currfile.FullName.IndexOf(".csproj", comp_ic) != currfile.FullName.Length - 7)
//                {
//                    continue;
//                }
//#if DEBUG
//                if (currfile.Name == $"{System.Reflection.Assembly.GetExecutingAssembly().GetName().Name}.csproj")
//                {
//                    continue; // don't include this project while testing
//                }
//#endif
//                filelines = File.ReadAllLines(currfile.FullName, Encoding.UTF8);
//                string assemblyname = "";
//                foreach (string currline in filelines)
//                {
//                    if (currline.Trim().StartsWith("'") |
//                        currline.Trim().StartsWith("//"))
//                    {
//                        continue;
//                    }
//                    // get the assembly name of the project
//                    if (currline.IndexOf("<AssemblyName>", comp_ic) >= 0)
//                    {
//                        assemblyname = Functions.SubstringBetween(currline, "<AssemblyName>", "<");
//                    }
//                }
//                if (versionlist[assemblyname] == origversionlist[assemblyname])
//                {
//                    continue;
//                }
//                // this project needs changes
//                Console.WriteLine($"Saving {currfile.Name}...");
//                savecount += 1;
//                string newversion = versionlist[assemblyname];
//                bool changed = false;
//                StringBuilder result = new StringBuilder();
//                foreach (string currline in filelines)
//                {
//                    if (result.Length > 0)
//                        result.AppendLine();
//                    if (currline.Trim().StartsWith("'") |
//                        currline.Trim().StartsWith("//"))
//                    {
//                        result.Append(currline);
//                        continue;
//                    }
//                    // get the version number of the project
//                    if (currline.IndexOf("<ApplicationVersion>", comp_ic) >= 0)
//                    {
//                        result.Append(Functions.ReplaceBetween(currline, "<ApplicationVersion>", "<", newversion));
//                        changed = true;
//                        continue;
//                    }
//                    if (currline.IndexOf("<MinimumRequiredVersion>", comp_ic) >= 0)
//                    {
//                        result.Append(Functions.ReplaceBetween(currline, "<MinimumRequiredVersion>", "<", newversion));
//                        changed = true;
//                        continue;
//                    }
//                    if (currline.IndexOf("<ApplicationRevision>", comp_ic) >= 0)
//                    {
//                        // just get the last revision number
//                        result.Append(Functions.ReplaceBetween(currline, "<ApplicationRevision>", "<",
//                                                               newversion.Substring(newversion.LastIndexOf(".") + 1)));
//                        changed = true;
//                        continue;
//                    }
//                    // no change needed
//                    result.Append(currline);
//                }
//                check if anything changed
//                if (changed)
//                {
//                    // write out the changes
//                    if (File.Exists(currfile.FullName))
//                    {
//                        File.SetAttributes(currfile.FullName, FileAttributes.Normal);
//                        File.Delete(currfile.FullName);
//                    }
//                    File.WriteAllText(currfile.FullName, result.ToString());
//                }
//                // now update the version in the AssemblyInfo file
//                if (currfile.FullName.IndexOf(".vbproj", comp_ic) >= 0)
//                {
//                    UpdateAssemblyInfoVersion(currfile.DirectoryName +
//                                              "\\My Project\\AssemblyInfo.vb",
//                                              newversion);
//                }
//                else
//                {
//                    UpdateAssemblyInfoVersion(currfile.DirectoryName +
//                                              "\\Properties\\AssemblyInfo.cs",
//                                              newversion);
//                }
//            }
//            // loop through all the directories
//            foreach (DirectoryInfo tempdir in thisdir.GetDirectories())
//            {
//                if (tempdir.Name.StartsWith("."))
//                {
//                    continue;
//                }
//                savecount += SaveProgramVersions(tempdir);
//            }
//            return savecount;
//        }
    }
}
