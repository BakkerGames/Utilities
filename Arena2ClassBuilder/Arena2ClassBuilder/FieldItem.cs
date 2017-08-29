// FieldItem.cs - 08/29/2017

class FieldItem
{
    public string FieldName { get; set; }
    public string FieldType { get; set; }
    public string FieldLen { get; set; }
    public bool NotNull { get; set; }
    public bool IsIdentity { get; set; }
    public bool IsTimestamp { get; set; }
}
