// Programs.BatchFiles.cs - 03/05/2017

// 03/05/2017 - SBakker
//            - Ignore any directories starting with "." when clearing old object files.
// 01/25/2017 - SBakker
//            - Added handling for SourceFlag, where there is a "Source" directory next to "Bin".
// 09/20/2016 - SBakker
//            - Use "delims=;" in "for" batch command, so that paths with spaces work.
// 08/23/2016 - SBakker
//            - Simplified log file handling, made errors easier to find.
// 08/17/2016 - SBakker
//            - Remove each project's bin and obj directories after building. Yay!
// 08/12/2016 - SBakker
//            - Added extra messages, and include non-zero Warning(s).
// 08/09/2016 - SBakker
//            - Changed use of DevEnv to MSBuild. Muuuuuch faster!
// 08/02/2016 - SBakker
//            - Fixed syntax of "findstr" to give correct results.
// 07/29/2016 - SBakker
//            - Always ignore "Test_" projects when creating BuildAll.bat

using System.IO;
using System.Text;

namespace UpdateVersions2
{
    partial class Program
    {

        //static string buildprogname = "\"C:\\Program Files (x86)\\Microsoft Visual Studio 14.0\\Common7\\IDE\\devenv.exe\"";
        static string buildprogname = "\"C:\\Program Files (x86)\\MSBuild\\14.0\\Bin\\MSBuild.exe\"";
        static string buildopts = "/p:Configuration=Release /clp:ErrorsOnly /verbosity:Normal /NoLogo";

        private static void BuildBatchFiles(DirectoryInfo thisdir)
        {
            int currlevel = 0;
            bool anyfound;
            StringBuilder result = new StringBuilder();
            result.AppendLine("@echo off");
            result.AppendLine($"set buildprog={buildprogname}");
            result.AppendLine($"set buildopts={buildopts}");
            result.AppendLine($"if not exist %buildprog% set buildprog={buildprogname.Replace(" (x86)", "")}");
            result.AppendLine("if not exist %buildprog% (");
            result.AppendLine("echo MSBuild compiler not found:");
            result.AppendLine("echo %buildprog%");
            result.AppendLine("echo.");
            result.AppendLine("pause");
            result.AppendLine("goto :eof");
            result.AppendLine(")");
            result.AppendLine();
            result.AppendLine("echo --- Clearing old object files ---");
            result.AppendLine();
            result.AppendLine("del _delbin.txt >nul 2>nul");
            result.AppendLine("dir /ad /s /b . | find \"\\bin\" | find /v \"\\bin\\\" | find /v \"\\.\" >>_delbin.txt");
            result.AppendLine("dir /ad /s /b . | find \"\\obj\" | find /v \"\\obj\\\" | find /v \"\\.\" >>_delbin.txt");
            result.AppendLine("for /f \"delims=;\" %%a in (_delbin.txt) do rmdir /s /q \"%%a\"");
            result.AppendLine("del _delbin.txt >nul 2>nul");
            result.AppendLine();
            if (SourceFlag)
            {
                result.AppendLine("del ..\\Bin\\*.exe >nul 2>nul");
                result.AppendLine("del ..\\Bin\\*.dll >nul 2>nul");
                result.AppendLine("del ..\\Bin\\*.config >nul 2>nul");
                result.AppendLine("del ..\\Bin\\*.settings >nul 2>nul");
            }
            else
            {
                result.AppendLine("rename \"Bin\\Arena.xml\" \"Arena_xml.delbin\" >nul 2>nul");
                result.AppendLine("del Bin\\*.exe >nul 2>nul");
                result.AppendLine("del Bin\\*.dll >nul 2>nul");
                result.AppendLine("del Bin\\*.xml >nul 2>nul");
                result.AppendLine("del Bin\\*.config >nul 2>nul");
                result.AppendLine("del Bin\\*.settings >nul 2>nul");
                result.AppendLine("rename \"Bin\\Arena_xml.delbin\" \"Arena.xml\" >nul 2>nul");
            }
            result.AppendLine();
            result.AppendLine("set logfile=\"BuildAll.log\"");
            result.AppendLine("attrib -r %logfile% >nul 2>nul");
            result.AppendLine("del %logfile% >nul 2>nul");
            result.AppendLine();
            result.AppendLine("@echo on");
            result.AppendLine();
            do
            {
                anyfound = false;
                foreach (string currproj in levellist.Keys)
                {
                    if (levellist[currproj] == currlevel)
                    {
                        if (currproj.IndexOf("Test_", comp_ic) == 0)
                        {
                            continue; // Always ignore Test_ projects here
                        }
                        if (!anyfound)
                        {
                            result.AppendLine("REM");
                            result.AppendLine($"REM --- Level {currlevel} ---");
                            anyfound = true;
                        }
                        string projectname = projectlist[currproj].Substring(thisdir.FullName.Length + 1);
                        if (!File.Exists($"{thisdir.FullName}\\{projectname}"))
                        {
                            result.AppendLine($"@echo Not found! {thisdir.FullName}\\{projectname}");
                        }
                        else
                        {
                            //result.AppendLine($"@echo {thisdir.FullName}\\{projectname} >>%logfile%");
                            result.AppendLine($"%buildprog% \"{projectname}\" %buildopts% >>%logfile%");
                        }
                    }
                }
                currlevel += 1;
                if (anyfound)
                {
                    result.AppendLine();
                }
            } while (anyfound);
            result.AppendLine("@echo off");
            result.AppendLine("del _delbin.txt >nul 2>nul");
            result.AppendLine("dir /ad /s /b . | find \"\\bin\" | find /v \"\\bin\\\" | find /v \"\\.\" >>_delbin.txt");
            result.AppendLine("dir /ad /s /b . | find \"\\obj\" | find /v \"\\obj\\\" | find /v \"\\.\" >>_delbin.txt");
            result.AppendLine("for /f %%a in (_delbin.txt) do rmdir /s /q \"%%a\"");
            result.AppendLine("del _delbin.txt >nul 2>nul");
            result.AppendLine();
            result.AppendLine("for %%F in (%logfile%) do if %%~zF equ 0 del \"%%F\"");
            result.AppendLine("if not exist %logfile% goto :noerrors");
            result.AppendLine();
            result.AppendLine("@echo on");
            result.AppendLine("REM");
            result.AppendLine("REM --- Errors ---");
            result.AppendLine("@echo.");
            result.AppendLine("@more %logfile%");
            result.AppendLine("@goto :pause");
            result.AppendLine();
            result.AppendLine(":noerrors");
            result.AppendLine("@echo.");
            result.AppendLine("@echo --- No Errors Found ---");
            result.AppendLine();
            result.AppendLine(":pause");
            result.AppendLine("@echo.");
            result.AppendLine("@pause");
            result.AppendLine();
            result.AppendLine(":eof");
            // write out the batch file
            File.WriteAllText($"{thisdir.FullName}\\BuildAll.bat", result.ToString());
        }

    }
}
