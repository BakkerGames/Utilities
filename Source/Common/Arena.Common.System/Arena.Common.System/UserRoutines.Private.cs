// UserRoutines.Private.cs - 12/21/2017

namespace Arena.Common.System
{
    public static partial class UserRoutines
    {
        private static void SetLoginID_IDRIS(string userName)
        {
            _loginID_IDRIS = userName.ToUpper();
            if (_loginID_IDRIS.Length > 5)
            {
                _loginID_IDRIS = _loginID_IDRIS.Substring(0, 5);
            }
        }
    }
}
