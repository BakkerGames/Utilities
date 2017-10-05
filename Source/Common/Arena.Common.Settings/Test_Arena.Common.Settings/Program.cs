// Program.cs - 09/30/2017

using Arena.Common.Settings;
using Arena.Common.System;
using System;

namespace Test_Arena.Common.Settings
{
    class Program
    {
        static void Main(string[] args)
        {
            ShowProductFamily(ProductFamily.Arena);
            ShowProductFamily(ProductFamily.Arena2);
            ShowProductFamily(ProductFamily.IDRIS);
            ShowProductFamily(ProductFamily.Advantage);
            ShowProductFamily(ProductFamily.Common);
            ShowProductFamily(ProductFamily.Security);
            Console.Write("Press any key to continue...");
            Console.ReadKey();
        }

        static void ShowProductFamily(string productFamily)
        {
            Console.WriteLine($"{productFamily}:");
            string ArenaServer = DataSettings.GetServerName(productFamily);
            Console.WriteLine($"    Server Name   = {ArenaServer}");
            string ArenaDatabase = DataSettings.GetDatabaseName(productFamily);
            Console.WriteLine($"    Database Name = {ArenaDatabase}");
            Console.WriteLine();
        }
    }
}
