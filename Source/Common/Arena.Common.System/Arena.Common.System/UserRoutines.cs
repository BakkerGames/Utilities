// UserRoutines.cs - 12/21/2017

using Arena.Common.Errors;
using System;

namespace Arena.Common.System
{
    public static partial class UserRoutines
    {
        private static string _loginID;
        private static string _loginID_IDRIS;
        private static string _userDomainName;
        private static string _userFullName;

        static UserRoutines()
        {
            ResetUser();
        }

        public static void ResetUser()
        {
            _loginID = Environment.UserName.ToLower();
            _userFullName = _loginID;
            SetLoginID_IDRIS(_loginID);
            _userDomainName = Environment.UserDomainName.ToLower();
        }

        public static string GetLoginID()
        {
            return _loginID;
        }

        public static string GetLoginID_IDRIS()
        {
            return _loginID_IDRIS;
        }

        public static string GetUserDomainName()
        {
            return _userDomainName;
        }

        public static void SetUserFullName(string userFullName)
        {
            // user full name is not available at this level, so must be set from above
            if (string.IsNullOrEmpty(userFullName))
            {
                throw new SystemException(ErrorHandler.FixMessage("UserFullName cannot be blank"));
            }
            _userFullName = userFullName;
        }

        public static void ImpersonateLoginID(string loginID)
        {
            ImpersonateLoginID(loginID, null);
        }

        public static void ImpersonateLoginID(string loginID, string userDomainName)
        {
            if (string.IsNullOrEmpty(loginID))
            {
                throw new SystemException(ErrorHandler.FixMessage("LoginID cannot be blank"));
            }
            _loginID = loginID.Trim().ToLower();
            _userFullName = _loginID;
            SetLoginID_IDRIS(_loginID);
            if (!string.IsNullOrEmpty(userDomainName?.Trim()))
            {
                _userDomainName = userDomainName.Trim().ToLower();
            }
        }
    }
}
