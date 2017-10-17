// ErrorHandler.cs - 10/17/2017

using System;

namespace Arena.Common.Errors
{
    public static class ErrorHandler
    {
        public static string FixMessage(string message)
        {
            return $"{message}\r\n\r\nCall Stack:\r\n{Environment.StackTrace}";
        }

        public static string FixMessage(string message, string internalMessage)
        {
            return $"{message}\r\n\r\n{internalMessage}\r\n\r\nCall Stack:\r\n{Environment.StackTrace}";
        }
    }
}
