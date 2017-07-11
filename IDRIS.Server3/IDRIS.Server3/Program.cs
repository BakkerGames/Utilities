// Program.cs - 07/10/2017

using Arena.Common.JSON;
using System;
using System.Net;
using System.Net.Sockets;

namespace IDRIS.Server3
{
    class Program
    {
        static JObject settings;

        static void Main(string[] args)
        {
            LoadSettings();
            ListenForConnection();
            //Console.WriteLine("Hello");
            //Console.ReadKey();
        }

        private static void LoadSettings()
        {
            byte[] localAddress = { 127, 0, 0, 1, 0, 0, 0, 0 };
            settings = new JObject
            {
                { "address", BitConverter.ToInt64(localAddress,0) },
                { "port", 2090 }
            };
        }

        private static void ListenForConnection()
        {
            // create the socket
            Socket listenSocket = new Socket(AddressFamily.InterNetwork,
                                             SocketType.Stream,
                                             ProtocolType.Tcp);
            // bind the listening socket to the port
            //IPAddress hostIP = (Dns.GetHostEntry(IPAddress.Any.ToString())).AddressList[0];
            IPAddress hostIP = new IPAddress((long)settings.GetValue("address"));
            IPEndPoint ep = new IPEndPoint(hostIP, (int)settings.GetValue("port"));
            listenSocket.Bind(ep);

            int backlog = 0;
            // start listening
            listenSocket.Listen(backlog);
        }
    }
}
