// FieldItem.cs - 02/28/2018

class FieldItem
{
    public string FieldName { get; set; }
    public string SQLFieldName { get; set; }
    public string FieldType { get; set; }
    public string FieldLen { get; set; }
    public bool NotNull { get; set; }
    public bool IsIdentity { get; set; }
    public bool IsTimestamp { get; set; }
}
