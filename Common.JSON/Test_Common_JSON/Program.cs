using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Common.JSON;

namespace Test_Common_JSON
{
    class Program
    {
        static void Main(string[] args)
        {
            JSONObject jo = new JSONObject("{\r\n  \"key\": 123,\r\n  \"otherkey\": 789.12\r\n}");
            Console.WriteLine(jo.ToString());
            Console.Read();
        }
    }
}
