// Programs.Find.cs - 10/19/2016

// 10/19/2016 - SBakker
//            - Skip directory names starting with "_".
// 09/30/2016 - SBakker
//            - Added .sqlproj support.
// 08/29/2016 - SBakker
//            - Fixed so references with extra info are gracefully used, not an error.
// 08/17/2016 - SBakker
//            - Fixing situation where project reference has Version and other info embedded.
// 08/16/2016 - SBakker
//            - Removed the saving of version numbers. They have no meaning when using Git.
// 07/29/2016 - SBakker
//            - Don't exclude "Test_" projects during Find.
//            - Skip directory names starting with "."
//            - Added error checking during Find.

using System;
using System.IO;
using System.Collections.Generic;
using System.Text;

namespace UpdateVersions2
{
    partial class Program
    {
        private static int FindProgramVersions(DirectoryInfo thisdir)
        {
            string[] filelines;
            // loop through all the files
            foreach (FileInfo currfile in thisdir.GetFiles())
            {
                if (currfile.FullName.IndexOf(".vbproj", comp_ic) != currfile.FullName.Length - 7 &
                    currfile.FullName.IndexOf(".csproj", comp_ic) != currfile.FullName.Length - 7 &
                    currfile.FullName.IndexOf(".sqlproj", comp_ic) != currfile.FullName.Length - 8)
                {
                    continue;
                }
#if DEBUG
                if (currfile.Name == $"{System.Reflection.Assembly.GetExecutingAssembly().GetName().Name}.csproj")
                {
                    continue; // don't include this project while testing
                }
#endif
                Console.WriteLine($"Checking {currfile.Name}...");
                filelines = File.ReadAllLines(currfile.FullName, Encoding.UTF8);
                List<string> subfiles = new List<string>();
                string assemblyname = "";
                string subfilename;
                string referencename;
                string version = "";
                // loop through all the lines of the project file
                foreach (string currline in filelines)
                {
                    if (currline.Trim().StartsWith("'") |
                        currline.Trim().StartsWith("//"))
                    {
                        continue;
                    }
                    // get the assembly name of the project
                    if (currline.IndexOf("<AssemblyName>", comp_ic) >= 0)
                    {
                        assemblyname = Functions.SubstringBetween(currline, "<AssemblyName>", "<");
                    }
                    // get the version number of the project
                    if (currline.IndexOf("<ApplicationVersion>", comp_ic) >= 0)
                    {
                        version = Functions.SubstringBetween(currline, "<ApplicationVersion>", "<");
                    }
                    // get the compile include filenames
                    if (currline.IndexOf("<Compile Include=\"", comp_ic) >= 0)
                    {
                        subfilename = Functions.SubstringBetween(currline, "<Compile Include=\"", "\"");
                        // have to exclude AssemblyInfo, as it won't have the right datetime
                        if (subfilename.IndexOf("AssemblyInfo.", comp_ic) < 0)
                        {
                            subfiles.Add(subfilename);
                        }
                    }
                    // get the embedded resource filenames
                    if (currline.IndexOf("<EmbeddedResource Include=\"", comp_ic) >= 0)
                    {
                        subfilename = Functions.SubstringBetween(currline, "<EmbeddedResource Include=\"", "\"");
                        // have to exclude AssemblyInfo, as it won't have the right datetime
                        if (subfilename.IndexOf("AssemblyInfo.", comp_ic) < 0)
                        {
                            subfiles.Add(subfilename);
                        }
                    }
                    // get the miscellaneous resource filenames
                    if (currline.IndexOf("<None Include=\"", comp_ic) >= 0)
                    {
                        subfilename = Functions.SubstringBetween(currline, "<None Include=\"", "\"");
                        // have to exclude AssemblyInfo, as it won't have the right datetime
                        if (subfilename.IndexOf("AssemblyInfo.", comp_ic) < 0)
                        {
                            subfiles.Add(subfilename);
                        }
                    }
                }
                foreach (string currline in filelines)
                {
                    // now get the references, which might have newer version numbers
                    if (currline.IndexOf("<Reference Include=\"", comp_ic) >= 0)
                    {
                        referencename = Functions.SubstringBetween(currline, "<Reference Include=\"", "\"");
                        // ignore known system references
                        if (referencename == "System" |
                            referencename.StartsWith("System.") |
                            referencename.StartsWith("Microsoft."))
                        {
                            continue;
                        }
                        if (referencename.IndexOf(",") >= 0)
                        {
                            referencename = referencename.Substring(0, referencename.IndexOf(",")).Trim();
                            //throw new SystemException($"Invalid reference {referencename} in project {currfile.Name} - Please drop and re-add");
                        }
                        referencelist.Add($"{assemblyname}:{referencename}");
                    }
                }
                // some versions aren't in the project files, they are in the assemblyinfo files
                if (string.IsNullOrEmpty(version))
                {
                    if (currfile.FullName.IndexOf(".vbproj", comp_ic) >= 0)
                    {
                        version = GetAssemblyInfoVersion(currfile.DirectoryName +
                                                         "\\My Project\\AssemblyInfo.vb");
                    }
                    else
                    {
                        version = GetAssemblyInfoVersion(currfile.DirectoryName +
                                                         "\\Properties\\AssemblyInfo.cs");
                    }
                }
                string newverwsion = version;
                foreach (string currsubfile in subfiles)
                {
                    string subfilefullname = $"{currfile.DirectoryName}\\{currsubfile}";
                    if (!File.Exists(subfilefullname))
                    {
                        // not an error, could be a system reference
                        continue;
                    }
                    FileInfo subfileinfo = new FileInfo(subfilefullname);
                    string tempversion = subfileinfo.LastWriteTimeUtc.ToString(FileDateVersionFormat);
                    if (Functions.CompareVersions(newverwsion, tempversion) < 0)
                    {
                        newverwsion = tempversion;
                    }
                }
                if (projectlist.ContainsKey(assemblyname))
                {
                    Console.WriteLine();
                    Console.WriteLine($"ERROR: Duplicate assembly name found: {assemblyname}");
                    Console.WriteLine(projectlist[assemblyname]);
                    Console.WriteLine(currfile.FullName);
                    return -1;
                }
                projectlist.Add(assemblyname, currfile.FullName);
                //origversionlist.Add(assemblyname, version);
                //versionlist.Add(assemblyname, newverwsion);
                levellist.Add(assemblyname, 0);
            }
            // loop through all the directories
            foreach (DirectoryInfo tempdir in thisdir.GetDirectories())
            {
                if (tempdir.Name.StartsWith(".") || tempdir.Name.StartsWith("_"))
                {
                    continue;
                }
                int result = FindProgramVersions(tempdir);
                if (result != 0)
                {
                    return result;
                }
            }
            return 0;
        }
    }
}
