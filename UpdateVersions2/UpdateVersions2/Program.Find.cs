// Programs.Find.cs - 01/11/2018

// 01/11/2018 - SBakker
//            - Must return postive number for ERRORLEVEL to work.
// 11/10/2017 - SBakker
//            - Write out UTF8 without BOM.
// 10/12/2017 - SBakker
//            - Remove version, etc from <Reference Include> lines. Remove <SpecificVersion> lines.
//              This fixes a bug in adding references to the list, as well as making compares better.
//            - Throw error if SpecificVersion=True found.
//            - Show message for project files fixed.
// 08/29/2017 - SBakker
//            - Throw error if <ProjectReference> found. It is used in debugging and must
//              be fixed before compiling can happen.
// 04/27/2017 - SBakker
//            - Ignore any "Reference Include" values with a <HintPath> containing "\\packages\\".
//              These are NuGet packages installed for the specific project.
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
using System.Collections.Generic;
using System.IO;
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
                    if (currline.Trim().StartsWith("'") ||
                        currline.Trim().StartsWith("//"))
                    {
                        continue;
                    }
                    // check for project references
                    if (currline.IndexOf("<ProjectReference", comp_ic) >= 0)
                    {
                        if (assemblyname.IndexOf("Test", comp_ic) != 0
                            && assemblyname.IndexOf("UnitTest", comp_ic) != 0)
                        {
                            Console.WriteLine();
                            Console.WriteLine($"ERROR: Debugging <ProjectReference> found: {assemblyname}");
                            Console.WriteLine(currfile.FullName);
                            return 1;
                        }
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
                referencename = null;
                bool checkNextLine = false;
                string newCurrLine = "";
                bool projectChanged = false;
                StringBuilder newProjectFile = new StringBuilder();
                foreach (string currline in filelines)
                {
                    newCurrLine = currline;
                    if (checkNextLine && referencename != null)
                    {
                        if (currline.IndexOf("<SpecificVersion", comp_ic) >= 0)
                        {
                            if (currline.IndexOf(">True<", comp_ic) >= 0)
                            {
                                Console.WriteLine();
                                Console.WriteLine($"ERROR: SpecificVersion=True found: {assemblyname}");
                                Console.WriteLine(currfile.FullName);
                                return 1;
                            }
                            projectChanged = true;
                            continue; // don't include in newProjectFile
                        }
                        if (currline.IndexOf("<HintPath>", comp_ic) >= 0 &&
                            currline.IndexOf("\\packages\\", comp_ic) >= 0)
                        {
                            referencename = null;
                        }
                        if (referencename != null)
                        {
                            referencelist.Add($"{assemblyname}:{referencename}");
                            referencename = null;
                        }
                    }
                    referencename = null;
                    checkNextLine = false;
                    // now get the references, which might have newer version numbers
                    if (currline.IndexOf("<Reference Include=\"", comp_ic) >= 0)
                    {
                        referencename = Functions.SubstringBetween(currline, "<Reference Include=\"", "\"");
                        // ignore known system references
                        if (referencename == "System" ||
                            referencename.StartsWith("System.") ||
                            referencename.StartsWith("Microsoft."))
                        {
                            referencename = null;
                        }
                        if (!string.IsNullOrEmpty(referencename))
                        {
                            if (referencename.IndexOf(",") >= 0)
                            {
                                referencename = referencename.Substring(0, referencename.IndexOf(",")).Trim();
                                newCurrLine = $"{currline.Substring(0, currline.IndexOf("\"") + 1)}{referencename}\">";
                                projectChanged = true;
                            }
                            checkNextLine = true;
                        }
                    }
                    if (newProjectFile.Length > 0)
                    {
                        newProjectFile.AppendLine();
                    }
                    newProjectFile.Append(newCurrLine); // so last line has no crlf
                }
                if (projectChanged)
                {
                    // project file needs changing
                    File.WriteAllText(currfile.FullName, newProjectFile.ToString(), new UTF8Encoding(false, true));
                    Console.WriteLine($"Updated file {currfile.FullName}");
                }
                if (referencename != null)
                {
                    referencelist.Add($"{assemblyname}:{referencename}");
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
                    return 1;
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
