using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

public partial class DynamicTemplate
{
    public DynamicTemplate()
    {
        this.DynamicTemplateAttributes = new HashSet<DynamicTemplateAttribute>();
    }

    [Key]
    public int Id { get; set; }
    public string Name { get; set; }
    public string Description { get; set; }

    public virtual ICollection<DynamicTemplateAttribute> DynamicTemplateAttributes { get; set; }

}