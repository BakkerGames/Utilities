// UserInfo.cs - 09/12/2017

using System;

namespace Arena.Common.System
{
    public static class UserInfo
    {
        private static string _username;
        private static string _username_IDRIS;

        static UserInfo()
        {
            _username = Environment.UserName.ToLower();
            _username_IDRIS = _username.ToUpper();
			if (_username_IDRIS.Length > 5)
			{
				_username_IDRIS = _username_IDRIS.Substring(0, 5);
			}
        }

        public static string GetLoginID()
        {
            return _username;
        }

        public static string GetLoginID_IDRIS()
        {
            return _username_IDRIS;
        }
    }
}
