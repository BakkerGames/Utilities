﻿// BootstrapAppConfig.cs - 09/27/2017

// ----------------------------------------------------------------------------------------------------------
// 09/27/2017 - SBakker - URD 15244
//            - Added CopyRecursive property.
// 09/21/2017 - SBakker - URD 15244
//            - Created BootstrapAppConfig class for holding AppConfig info.
// ----------------------------------------------------------------------------------------------------------

using System.Collections.Generic;

namespace Arena.Common.Bootstrap
{
    internal class BootstrapAppConfig
    {
        public string FullLaunchPath { get; set; }
        public List<string> AppPaths { get; set; }
        public bool? CopyRecursive { get; set; }
        public List<string> OtherAppPaths { get; set; }
    }
}
