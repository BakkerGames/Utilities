// UserInfo.cs - 10/13/2017

using Arena.Common.Errors;
using System;

namespace Arena.Common.System
{
    public static class UserInfo
    {
        private static string _userName;
        private static string _userName_IDRIS;
        private static string _domainName;
        private static string _userFullName;

        static UserInfo()
        {
            _userName = Environment.UserName.ToLower();
            _domainName = Environment.UserDomainName.ToLower();
            SetUserNameIDRIS(_userName);
            _userFullName = _userName;
        }

        private static void SetUserNameIDRIS(string userName)
        {
            _userName_IDRIS = userName.ToUpper();
            if (_userName_IDRIS.Length > 5)
            {
                _userName_IDRIS = _userName_IDRIS.Substring(0, 5);
            }
        }

        public static string GetLoginID()
        {
            return _userName;
        }

        public static string GetLoginID_IDRIS()
        {
            return _userName_IDRIS;
        }

        public static string GetUserDomainName()
        {
            return _domainName;
        }

        public static void SetUserFullName(string userFullName)
        {
            // this information is not available at this level, so must be set from above
            if (string.IsNullOrEmpty(userFullName))
            {
                throw new SystemException(ErrorHandler.FixMessage("UserFullName cannot be blank"));
            }
            _userFullName = userFullName;
        }

        public static void SetUserNameForWebServices(string userName)
        {
            if (string.IsNullOrEmpty(userName))
            {
                throw new SystemException(ErrorHandler.FixMessage("UserName cannot be blank"));
            }
            _userName = userName;
            _userFullName = userName;
            SetUserNameIDRIS(_userName);
        }

        public static void SetUserDomainNameForWebServices(string domainName)
        {
            if (string.IsNullOrEmpty(domainName))
            {
                throw new SystemException(ErrorHandler.FixMessage("DomainName cannot be blank"));
            }
            _domainName = domainName;
        }
    }
}
