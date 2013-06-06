using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Design;
using System.Windows.Forms;
using MobileDesigner.Controls;

namespace MobileDesigner.Design
{   
    public class LabelProperties 
    {
        public const string EmptyLabelText = "(blank)";

        protected readonly Label Control;

        public LabelProperties(Label control)
        {
            Control = control;
        }

        [Category("General")]
        public string Caption
        {
            get { return Control.Text == EmptyLabelText ? string.Empty : Control.Text; }
            set { Control.Text = string.IsNullOrWhiteSpace(value) ? EmptyLabelText : value; }
        }

        [Category("Appearance")]
        [TypeConverter(typeof(FontConverter))]
        public Font Font
        {
            get { return Control.Font; }
            set { Control.Font = value; }
        }

        [Category("Appearance")]
        [DisplayName("Fore Colour")]
        public Color ForeColor
        {
            get { return Control.ForeColor; }
            set { Control.ForeColor = value; }
        }

        public virtual bool ShouldSerializeCaption()
        {
            return false;
        }
        public virtual bool ShouldSerializeFont()
        {
            return false;
        }
        public virtual bool ShouldSerializeForeColor()
        {
            return false;
        }
    }

    public class TextBoxProperties
    {
        protected readonly TextBox Control;

        public TextBoxProperties(TextBox control)
        {
            Control = control;
        }

        [Category("Appearance")]
        [TypeConverter(typeof(FontConverter))]
        public Font Font
        {
            get { return Control.Font; }
            set { Control.Font = value; }
        }

        [Category("Appearance")]
        [DisplayName("Fore Colour")]
        public Color ForeColor
        {
            get { return Control.ForeColor; }
            set { Control.ForeColor = value; }
        }

        public virtual bool ShouldSerializeFont()
        {
            return false;
        }
        public virtual bool ShouldSerializeForeColor()
        {
            return false;
        }
    }


    public class IconButtonProperties
    {
        protected readonly IconButton Control;

        public IconButtonProperties(IconButton control)
        {
            Control = control;
        }

        [Category("General")]
        public string Caption
        {
            get { return Control.Caption; }
            set { Control.Caption = value; }
        }

        [Category("Appearance")]
        [DisplayName("Fore Colour")]
        public Color ForeColor
        {
            get { return Control.CaptionColor; }
            set { Control.CaptionColor = value; }
        }

        [Category("Appearance")]
        [Editor(typeof(PictureEditor), typeof(UITypeEditor))]
        [TypeConverter(typeof(PictureConverter))]
        public Picture Image
        {
            get { return Control.Image; }
            set { Control.Image = value; }
        }

        public virtual bool ShouldSerializeCaption()
        {
            return false;
        }
        public virtual bool ShouldSerializeForeColor()
        {
            return false;
        }
        public virtual bool ShouldSerializeImage()
        {
            return false;
        }
    }

    public class HeaderPanelProperties
    {
        protected readonly HeaderPanel Control;

        public HeaderPanelProperties(HeaderPanel control)
        {
            Control = control;
        }

        [Category("Appearance")]
        [DisplayName("Back Colour")]
        public Color BackgroundColour
        {
            get { return Control.BackColor; }
            set { Control.BackColor = value; }
        }

        [Category("Appearance")]
        [DisplayName("Back Image")]
        [Editor(typeof(PictureEditor), typeof(UITypeEditor))]
        [TypeConverter(typeof(PictureConverter))]
        public Picture BackgroundImage
        {
            get { return Control.BackgroundImage; }
            set { Control.BackgroundImage = value; }
        }

        [Category("Appearance")]
        [DisplayName("Back Image Layout")]
        public ImageLayout BackgroundImageLayout
        {
            get { return Control.BackgroundImageLayout; }
            set { Control.BackgroundImageLayout = value; }
        }

        [Category("Appearance")]
        [DisplayName("Logo Image")]
        [Editor(typeof(PictureEditor), typeof(UITypeEditor))]
        [TypeConverter(typeof(PictureConverter))]
        public Picture LogoImage
        {
            get { return Control.LogoImage; }
            set { Control.LogoImage = value; }
        }

        [Category("Appearance")]
        [DisplayName("Logo Image Size")]
        public Size LogoImageSize
        {
            get { return Control.LogoImageSize; }
            set { Control.LogoImageSize = value; }
        }

        [Category("Appearance")]
        [DisplayName("Logo Image Align")]
        public ImageLocation LogoImageLocation
        {
            get { return Control.LogoImageLocation; }
            set { Control.LogoImageLocation = value; }
        }

        [Category("Appearance")]
        [DisplayName("Logo Image Offset")]
        public Point LogoImageOffset
        {
            get { return Control.LogoImageOffset; }
            set { Control.LogoImageOffset = value; }
        }

        public virtual bool ShouldSerializeBackgroundColour()
        {
            return false;
        }
        public virtual bool ShouldSerializeBackgroundImage()
        {
            return false;
        }
        public virtual bool ShouldSerializeBackgroundImageLayout()
        {
            return false;
        }
        public virtual bool ShouldSerializeLogoImage()
        {
            return false;
        }
        public virtual bool ShouldSerializeLogoImageSize()
        {
            return false;
        }
        public virtual bool ShouldSerializeLogoImageLocation()
        {
            return false;
        }
        public virtual bool ShouldSerializeLogoImageOffset()
        {
            return false;
        }
    }

    public class FooterPanelProperties
    {
        protected readonly FooterPanel Control;

        public FooterPanelProperties(FooterPanel control)
        {
            Control = control;
        }

        [Category("Appearance")]
        [DisplayName("Back Colour")]
        public Color BackgroundColour
        {
            get { return Control.BackColor; }
            set { Control.BackColor = value; }
        }

        [Category("Appearance")]
        [DisplayName("Back Image")]
        [Editor(typeof(PictureEditor), typeof(UITypeEditor))]
        [TypeConverter(typeof(PictureConverter))]
        public Picture BackgroundImage
        {
            get { return Control.BackgroundImage; }
            set { Control.BackgroundImage = value; }
        }

        [Category("Appearance")]
        [DisplayName("Back Image Layout")]
        public ImageLayout BackgroundImageLayout
        {
            get { return Control.BackgroundImageLayout; }
            set { Control.BackgroundImageLayout = value; }
        }

        public virtual bool ShouldSerializeBackgroundColour()
        {
            return false;
        }
        public virtual bool ShouldSerializeBackgroundImage()
        {
            return false;
        }
        public virtual bool ShouldSerializeBackgroundImageLayout()
        {
            return false;
        }
    }

    public class MainPanelProperties
    {
        protected readonly MainPanel Control;

        public MainPanelProperties(MainPanel control)
        {
            Control = control;
        }

        [Category("Appearance")]
        [DisplayName("Back Colour")]
        public Color BackgroundColour
        {
            get { return Control.BackColor; }
            set { Control.BackColor = value; }
        }

        [Category("Appearance")]
        [DisplayName("Back Image")]
        [Editor(typeof(PictureEditor), typeof(UITypeEditor))]
        [TypeConverter(typeof(PictureConverter))]
        public Picture BackgroundImage
        {
            get { return Control.BackgroundImage; }
            set { Control.BackgroundImage = value; }
        }

        [Category("Appearance")]
        [DisplayName("Back Image Layout")]
        public ImageLayout BackgroundImageLayout
        {
            get { return Control.BackgroundImageLayout; }
            set { Control.BackgroundImageLayout = value; }
        }

        public virtual bool ShouldSerializeFont()
        {
            return false;
        }
        public virtual bool ShouldSerializeForeColor()
        {
            return false;
        }
        public virtual bool ShouldSerializeBackgroundColour()
        {
            return false;
        }
        public virtual bool ShouldSerializeBackgroundImage()
        {
            return false;
        }
        public virtual bool ShouldSerializeBackgroundImageLayout()
        {
            return false;
        }
    }

    public class LinkButtonProperties
    {
        protected readonly LinkButton Control;

        public LinkButtonProperties(LinkButton control)
        {
            Control = control;
        }

        [DisplayName("\tTitle Font")]
        [Category("Appearance")]
        [TypeConverter(typeof(FontConverter))]
        public Font TitleFont
        {
            get { return Control.TitleFont; }
            set { Control.TitleFont = value; }
        }

        [Category("Appearance")]
        [DisplayName("\tTitle Fore Colour")]
        public Color TitleForeColor
        {
            get { return Control.TitleForeColor; }
            set { Control.TitleForeColor = value; }
        }

        [DisplayName("Description Font")]
        [Category("Appearance")]
        [TypeConverter(typeof(FontConverter))]
        public Font DescriptionFont
        {
            get { return Control.DescriptionFont; }
            set { Control.DescriptionFont = value; }
        }

        [Category("Appearance")]
        [DisplayName("Description Fore Colour")]
        public Color DescriptionForeColor
        {
            get { return Control.DescriptionForeColor; }
            set { Control.DescriptionForeColor = value; }
        }

        public virtual bool ShouldSerializeTitleFont()
        {
            return false;
        }

        public virtual bool ShouldSerializeTitleForeColor()
        {
            return false;
        }

        public virtual bool ShouldSerializeDescriptionFont()
        {
            return false;
        }

        public virtual bool ShouldSerializeDescriptionForeColor()
        {
            return false;
        }
    }
    public class WorkflowItemsListProperties
    {
        protected readonly ItemsList Control;

        public WorkflowItemsListProperties(ItemsList control)
        {
            Control = control;
        }

        [DisplayName("Item Font")]
        [Category("Appearance")]
        [TypeConverter(typeof(FontConverter))]
        public Font ItemFont
        {
            get { return Control.ItemFont; }
            set { Control.ItemFont = value; }
        }

        [Category("Appearance")]
        [DisplayName("Item Fore Colour")]
        public Color ItemForeColor
        {
            get { return Control.ItemForeColor; }
            set { Control.ItemForeColor = value; }
        }

        [DisplayName("\tUser Group")]
        [Category("\tSelection")]
        [TypeConverter(typeof(UserGroupConverter))]
        public UserGroup SelectedUserGroup
        {
            get { return Control.SelectedUserGroup; }
            set { Control.SelectedUserGroup = value; }
        }

        [DisplayName("Items")]
        [Category("\tSelection")]
        [TypeConverter(typeof(ItemsConverter))]
        [Editor(typeof(ItemsEditor), typeof(UITypeEditor))]
        public IList<Workflow> SelectedWorkflows
        {
            get { return Control.SelectedWorkflows; }
            set { Control.SelectedWorkflows = value; }
        }

        [Browsable(false)]
        public IList<UserGroup> UserGroups
        {
            get { return Control.UserGroups; }
        }

        [Browsable(false)]
        public IList<Workflow> Workflows
        {
            get { return Control.Workflows; }
        }

        public virtual bool ShouldSerializeItemFont()
        {
            return false;
        }

        public virtual bool ShouldSerializeItemForeColor()
        {
            return false;
        }

        public virtual bool ShouldSerializeSelectedUserGroup()
        {
            return false;
        }

        public virtual bool ShouldSerializeSelectedWorkflows()
        {
            return false;
        }
    }

}
