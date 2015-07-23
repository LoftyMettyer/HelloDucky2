
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

public partial class DynamicAttribute
{
    public DynamicAttribute()
    {
        this.DynamicTemplateAttributes = new HashSet<DynamicTemplateAttribute>();
    }

    public int Id { get; set; }
    public string Name { get; set; }

    public virtual ICollection<DynamicTemplateAttribute> DynamicTemplateAttributes { get; set; }
}