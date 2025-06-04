public class FieldInfo
{
    public string EntityLogicalName { get; set; }
    public string EntityDisplayName { get; set; }
    public string FieldSchemaName { get; set; }
    public string FieldDisplayName { get; set; }
    public string DataType { get; set; }
    public string RequiredLevel { get; set; }
    public string Description { get; set; }
    public List<ScriptReference> ScriptReferences { get; set; } = new();
}
public class ScriptReference
{
    public string WebResourceName { get; set; }
    public string Note { get; set; }
}