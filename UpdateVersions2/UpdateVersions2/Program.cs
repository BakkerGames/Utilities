// Programs.cs - 01/25/2017

// 01/25/2017 - SBakker
//            - Added handling for SourceFlag, where there is a "Source" directory next to "Bin".
// 09/02/2016 - SBakker
//            - Adding better error handling.
// 08/17/2016 - SBakker
//            - Display thrown errors as console messages, from FindProgramVersions().
// 08/16/2016 - SBakker
//            - Removed the saving of version numbers. They have no meaning when using Git.
// 07/29/2016 - SBakker
//            - Added error checking during Find.
//            - Removed "testflag". "Test_" projects can be handled without it.
// 07/21/2016 - SBakker
//            - Added updatecount to UpdateReferenceVersions(), for blank line management.
// 07/20/2016 - SBakker
//            - Added savecount to SaveProgramVersions(), for blank line management.

using System;
using System.IO;
using System.Collections.Generic;

namespace UpdateVersions2
{
    partial class Program
    {
        // StringComparison for ignoring case
        private static StringComparison comp_ic = StringComparison.OrdinalIgnoreCase;

        private static string FileDateVersionFormat = "yyyy.M.d.Hmm";

        private static Dictionary<string, string> projectlist =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        private static Dictionary<string, int> levellist =
            new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        private static List<string> referencelist =
            new List<string>();

        private static bool SourceFlag = false;

        static int Main(string[] args)
        {
            string startingdir = "";
            bool printsyntax = false;
            int result;
            for (int argnum = 0; argnum < args.Length; argnum++)
            {
                if (args[argnum].StartsWith("/"))
                {
                    // set any parameters
                    if (string.Compare(args[argnum], "/source", true) == 0)
                    {
                        SourceFlag = true;
                    }
                    // if error then printsyntax = true;
                }
                else
                {
                    if (!string.IsNullOrEmpty(startingdir))
                    {
                        printsyntax = true;
                    }
                    else
                    {
                        startingdir = args[argnum];
                    }
                }
            }
            if (String.IsNullOrEmpty(startingdir))
            {
                printsyntax = true;
            }
            else if (!Directory.Exists(startingdir))
            {
                Console.WriteLine($"Directory not found: {startingdir}");
                printsyntax = true;
            }
            if (printsyntax)
            {
                Console.WriteLine();
                Console.WriteLine($"Syntax: {System.Reflection.Assembly.GetExecutingAssembly().GetName().Name}.exe [options] <startingdirectory>");
#if DEBUG
                Console.Write("Press any key...");
                Console.ReadKey();
#endif
                return 1;
            }
            try
            {
                DirectoryInfo thisdir = new DirectoryInfo(startingdir);
                Console.WriteLine("Finding project version numbers...");
                Console.WriteLine();
                result = FindProgramVersions(thisdir);
                if (result != 0)
                {
#if DEBUG
                    Console.Write("Press any key...");
                    Console.ReadKey();
#endif
                    return result;
                }
                Console.WriteLine();
                Console.WriteLine("Finding referenced version numbers...");
                Console.WriteLine();
                int updatecount = UpdateReferenceVersions();
                if (updatecount > 0)
                {
                    Console.WriteLine();
                }
                BuildBatchFiles(thisdir);
#if DEBUG
                Console.Write("Press any key...");
                Console.ReadKey();
#endif
                return 0;
            }
            catch (SystemException ex)
            {
                Console.WriteLine();
                Console.WriteLine(ex.Message);
#if DEBUG
                Console.Write("Press any key...");
                Console.ReadKey();
#endif
                return 1;
            }
        }
    }
}
