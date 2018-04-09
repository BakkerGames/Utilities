// ErrorHandler.cs - 04/02/2018

using System;

namespace Arena.Common.Errors
{
    public static class ErrorHandler
    {
        public const string CallStackLabel = "Call Stack:";

        public static string FixMessage(string message)
        {
            int positionCallStack = message.IndexOf(CallStackLabel);
            if (positionCallStack >= 0)
            {
                return message;
            }
            return $"{message}\r\n\r\n{CallStackLabel}\r\n{Environment.StackTrace}";
        }

        public static string FixMessage(string message, string internalMessage)
        {
            int positionCallStack = message.IndexOf(CallStackLabel);
            int positionCallStackInternal = internalMessage.IndexOf(CallStackLabel);
            if (positionCallStack >= 0 || positionCallStackInternal >= 0)
            {
                return $"{message}\r\n\r\n{internalMessage}";
            }
            return $"{message}\r\n\r\n{internalMessage}\r\n\r\n{CallStackLabel}\r\n{Environment.StackTrace}";
        }

        public static string GetMessageInfo(string message)
        {
            int positionCallStack = message.IndexOf(CallStackLabel);
            if (positionCallStack < 0)
            {
                return message;
            }
            int messageLength = positionCallStack;
            while (messageLength >= 2 && message[messageLength - 2] == '\r' && message[messageLength - 1] == '\n')
            {
                messageLength -= 2;
            }
            return message.Substring(0, messageLength);
        }

        public static string GetCallStackInfo(string message)
        {
            int positionCallStack = message.IndexOf(CallStackLabel);
            if (positionCallStack < 0)
            {
                return null;
            }
            positionCallStack += CallStackLabel.Length + 2;
            return message.Substring(positionCallStack);
        }
    }
}
