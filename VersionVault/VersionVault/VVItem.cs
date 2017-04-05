// VVItem.cs - 03/14/2017

namespace VersionVault
{
    public class VVItem
    {
        public VVItem()
        {
        }
        public VVItem(string value)
        {
            ItemName = value;
        }
        public string ItemName { get; set; }
        public override string ToString()
        {
            string result = "";
            if (!string.IsNullOrEmpty(ItemName) && ItemName.Length >= 15)
            {
                result = $"{ItemName.Substring(0,4)}-{ItemName.Substring(4,2)}-{ItemName.Substring(6,2)} {ItemName.Substring(9, 2)}:{ItemName.Substring(11, 2)}:{ItemName.Substring(13, 2)}";
            }
            return result;
        }
    }
}
