// Programs.Update.cs - 09/02/2016

// 09/02/2016 - SBakker
//            - Adding better error handling.
// 08/16/2016 - SBakker
//            - Removed the saving of version numbers. They have no meaning when using Git.
// 07/21/2016 - SBakker
//            - Fixed reference not found.
//            - Only display "Comparing" when there is a version difference.

using System;

namespace UpdateVersions2
{
    partial class Program
    {
        private static int UpdateReferenceVersions()
        {
            bool changed;
            string assemblyname;
            string referencename;
            int result = 0;
            do
            {
                changed = false;
                foreach (string combinedname in referencelist)
                {
                    assemblyname = combinedname.Substring(0, combinedname.IndexOf(":"));
                    referencename = combinedname.Substring(combinedname.IndexOf(":") + 1);
                    try
                    {
                        if (levellist[assemblyname] <= levellist[referencename])
                        {
                            levellist[assemblyname] = levellist[referencename] + 1;
                            changed = true;
                        }
                    }
                    catch
                    {
                        throw new SystemException($"Error: Project {referencename} referenced by {assemblyname} but not found");
                    }
                }
            } while (changed);
            return result;
        }
    }
}
