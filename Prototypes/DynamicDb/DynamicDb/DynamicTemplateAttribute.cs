using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

public partial class DynamicTemplateAttribute
{
    public int TemplateId { get; set; }
    public int AttributeId { get; set; }
    public string DisplayName { get; set; }
    public int TypeId { get; set; }

    [Key]
    public int Idx { get; set; }

    [ForeignKey("AttributeId")]
    public virtual DynamicAttribute DynamicAttribute { get; set; }

    [ForeignKey("TemplateId")]
    public virtual DynamicTemplate DynamicTemplate { get; set; }

    [ForeignKey("TypeId")]
    public virtual DynamicType DynamicType { get; set; }
}