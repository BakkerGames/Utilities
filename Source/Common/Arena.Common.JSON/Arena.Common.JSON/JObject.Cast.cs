// JObject.Cast.cs - 04/06/2018

// this is for casting routines, where casting using (type?)value doesn't work.
// for example, numeric objects may be the wrong type, int instead of long, and
// won't cast. it is a bit unfortunate, having to create functions like these.

using System;
using System.Collections.Generic;

namespace Arena.Common.JSON
{
    sealed public partial class JObject : IEnumerable<KeyValuePair<string, object>>
    {
        public long? GetValueOrNullAsLong(string name)
        {
            if (_data.ContainsKey(name) && _data[name] != null)
            {
                return Convert.ToInt64(_data[name]);
            }
            return null;
        }
    }
}
