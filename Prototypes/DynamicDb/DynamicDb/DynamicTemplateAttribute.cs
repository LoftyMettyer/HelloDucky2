
using System.ComponentModel.DataAnnotations;

public partial class DynamicTemplateAttribute
{
    public int TemplateId { get; set; }
    public int AttributeId { get; set; }
    public string DisplayName { get; set; }
    public int TypeId { get; set; }
    public int Id { get; set; }

    public virtual DynamicAttribute DynamicAttribute { get; set; }
    public virtual DynamicTemplate DynamicTemplate { get; set; }
    public virtual DynamicType DynamicType { get; set; }
}